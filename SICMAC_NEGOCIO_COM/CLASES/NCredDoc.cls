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

Public rsAnalista As ADODB.Recordset
Dim rsAnaProducto As ADODB.Recordset

Dim sTelefono As String
'Dim sConvenio As String
'Dim sDireccionTrabajo As String

Public Event ProgresoBarra(ByVal I As Long, ByVal lsTitulo As String, ByVal lsSubtitulo As String)
Public Event IniciaBarra(ByVal lnTotal As Long)
Public Event FinalizaBarra()

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
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim Total As Double
Dim Lineas As Integer
Dim llena As String
Dim I As Integer
Lineas = 7
Set Dg = New DGarantia
Set rs = New ADODB.Recordset
Set rs = Dg.RecuperaTitularGarantReal(Nrog)
Cabecera = ImprimeCabeceraGarantiaReal(psNomAge, psFecha, psHora, rs!Nombre, psCodCMAC)
Total = 0
'Detalle

sSQL = " Select( Select cDescripcion from ProductoConcepto where nPrdConceptoCod=MGD.nPrdConceptoCod) Concepto,"
sSQL = sSQL & " nMonto"
sSQL = sSQL & " from MovGarantDet MGD"
sSQL = sSQL & " Where MGD.nMovNro =" & pnMovG
    
    Set loConecta = New DConecta
        loConecta.AbreConexion
        Set rs = loConecta.CargaRecordSet(sSQL)
        loConecta.CierraConexion
  
cuerpo = Cabecera
If rs.BOF And rs.EOF Then

Else

    cuerpo = cuerpo & Chr(10)
    While Not rs.EOF
     If psCodCMAC = "102" Then
        cuerpo = cuerpo & ImpreFormat(rs!concepto, 23, 0) & Space(8) & ImpreFormat(rs!nMonto, 10, 2, True) & Chr(10)
        Lineas = Lineas + 1
        Total = Total + rs!nMonto
     Else
        cuerpo = cuerpo & ImpreFormat(rs!concepto, 23, 0) & Space(8) & ImpreFormat(rs!nMonto, 10, 2, True) & Space(6) & ImpreFormat(rs!concepto, 23, 0, False) & Space(10) & ImpreFormat(rs!nMonto, 10, 2, True) & Chr(10)
        Lineas = Lineas + 1
        Total = Total + rs!nMonto
     End If
    rs.MoveNext
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
Set rs = Nothing
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
        'vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
        'vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
        'vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
        'vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
        'vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
        'vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
        'vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita
        
        vsCadImpresion = vsCadImpresion & lsNegritaOn
        'vsCadImpresion = vsCadImpresion & lsSaltoLin
        'vsCadImpresion = vsCadImpresion & lsSaltoLin
        'vsCadImpresion = vsCadImpresion & lsSaltoLin
        If psCodCMAC = "102" Then
            vsCadImpresion = vsCadImpresion & ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & ImpreFormat(Mid(psTipoCredito, 1, 25), 51, 0), 53) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsNegritaOff
            vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
        Else
            vsCadImpresion = vsCadImpresion & ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & ImpreFormat(Mid(psTipoCredito, 1, 25), 51, 0), 50, 0) & Trim(ImpreFormat(psNomCmacAbrev & " - CREDITOS" & Space(5) & ImpreFormat(Mid(psTipoCredito, 1, 25), 51, 0), 53, 0)) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 51, 0) & Trim(ImpreFormat(Mid(psNomAge, 1, 22) & " - " & psMoneda & Space(3) & Trim(psNroTransac), 53, 0)) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & lsNegritaOff
            vsCadImpresion = vsCadImpresion & ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 51, 0) & Trim(ImpreFormat("Fecha : " & psFecha & Space(5) & "Hora : " & psHora, 53, 0)) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat(Mid(psPersNombre, 1, 50), 51, 0) & Trim(ImpreFormat(Mid(psPersNombre, 1, 50), 53, 0)) & lsSaltoLin
            vsCadImpresion = vsCadImpresion & ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 51, 0) & ImpreFormat("Credito : " & psCtaCod & Space(2) & "Cuota:" & Mid(psCuotasPagadas, 1, 15), 53, 0) & lsSaltoLin
        End If
        
        Exit Sub

ErrorImprimeCabeceraBoletaPago:
        Err.Raise Err.Number, "Error En Proceso ImprimeCabeceraBoletaPago", Err.Description
    
End Sub
Public Sub ImprimeBoletaExtorno(ByVal psDescrip As String, ByVal psCtaCod As String, ByVal psPersNombre As String, ByVal psNomAge As String, _
            ByVal psMoneda As String, ByVal psCuotasPagadas As String, ByVal psFecha As String, ByVal psHora As String, _
            ByVal psNroTransac As String, ByVal pnCapPagado As Double, _
            ByVal pnSaldoCap As Double, ByVal psCodUsu As String, ByVal sLpt As String, Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC, _
            Optional ByVal psCodUserOri As String)
            
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
            psFecha, psHora, psPersNombre, "", psCtaCod, "", psNomCmacAbrev, gsCodCMAC)
    
    'Cuerpo de la Boleta
If gsCodCMAC = "102" Then
'    vsCadImpresion = vsCadImpresion & lsNegritaOn
'    vsCadImpresion = vsCadImpresion & ImpreFormat(String(7, "-") & psDescrip & String(7, "-"), 40) & lsSaltoLin
'    '45 menos la longitud de su etiqueta
'    vsCadImpresion = vsCadImpresion & lsSaltoLin
'    vsCadImpresion = vsCadImpresion & "Monto  :" & Space(13) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
'    vsCadImpresion = vsCadImpresion & lsSaltoLin
'    vsCadImpresion = vsCadImpresion & String(40, "-") & lsSaltoLin
'    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
'    vsCadImpresion = vsCadImpresion & lsNegritaOff
'    vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
'    'vsCadImpresion = vsCadImpresion & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & psCodUsu, 53, 0) & ImpreFormat("Prox. Fecha Pago : " & Space(10) & Space(3) & psCodUsu, 53, 0) & lsSaltoLin
'    vsCadImpresion = vsCadImpresion & lsSaltoLin
'    Print #nFicSal, vsCadImpresion
'    'Print #nFicSal, lsSaltoLin
'    Print #nFicSal, ""
'    Print #nFicSal, ""
'    Print #nFicSal, ""
'    Close #nFicSal
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
    vsCadImpresion = vsCadImpresion & "Usuario Exto. :" & Space(7) & ImpreFormat(psCodUsu, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Usuario Exto. :" & Space(7) & ImpreFormat(psCodUsu, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Usuario Gene. :" & Space(7) & ImpreFormat(psCodUserOri, 12)
    vsCadImpresion = vsCadImpresion & Space(16) & "Usuario Exto. :" & Space(7) & ImpreFormat(psCodUserOri, 12) & lsSaltoLin
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
            Optional ByVal pbPagoCheque As Boolean = False, Optional ByVal psNumCheque As String = "", Optional ByVal psCodCMAC As String = "", _
            Optional ByVal pnITF As Double = 0#, Optional ByVal pnInteresDesagio As Double = 0, Optional pbOpeRecCMAC As Boolean = False, Optional pbCargoAuto As Boolean = False)
            
Dim nFicSal As Integer


    vsCadImpresion = ""
    nFicSal = FreeFile

    Open sLpt For Output As nFicSal
    
    
    'Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
    'Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
    'Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
    'Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
    'Print #nFicSal, Chr$(27) + Chr$(15) ' condensada
    'Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
    'Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
    
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(64)
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(50)    'espaciamiento lineas 1/6 pulg.
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(67) & Chr$(22)   'Longitud de página a 22 líneas'
    vsCadImpresion = vsCadImpresion & Chr$(27) & Chr$(77)    'Tamaño 10 cpi
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(107) + Chr$(0)      'Tipo de Letra Sans Serif
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(18)  ' cancela condensada
    vsCadImpresion = vsCadImpresion & Chr$(27) + Chr$(72)  ' desactiva negrita
    vsCadImpresion = vsCadImpresion & lsSaltoLin
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
        vsCadImpresion = vsCadImpresion & Space(26) & "Cheque  :" & Space(2) & ImpreFormat(psNumCheque, 12) & lsSaltoLin
    End If
    If pbPagoCheque Then
        If pbOpeRecCMAC = True Then
            vsCadImpresion = vsCadImpresion & String(10, "-") & "PAGO CHEQUE RECEP. " & String(10, "-")
            vsCadImpresion = vsCadImpresion & Space(16) & String(10, "-") & "PAGO CHEQUE RECEP. " & String(10, "-") & lsSaltoLin
        Else
            vsCadImpresion = vsCadImpresion & String(10, "-") & "PAGO CON CHEQUE    " & String(10, "-")
            vsCadImpresion = vsCadImpresion & Space(16) & String(10, "-") & "PAGO CON CHEQUE    " & String(10, "-") & lsSaltoLin
        End If
    Else
        If pbOpeRecCMAC = True Then
            vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO.EFEC.RECEPC. " & String(10, "-")
            vsCadImpresion = vsCadImpresion & Space(16) & String(10, "-") & " PAGO.EFEC.RECEPC. " & String(10, "-") & lsSaltoLin
        Else
            If pbCargoAuto Then
                vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO CARGO CUENTA " & String(10, "-")
                vsCadImpresion = vsCadImpresion & Space(16) & String(10, "-") & " PAGO CARGO CUENTA " & String(10, "-") & lsSaltoLin
            Else
                vsCadImpresion = vsCadImpresion & String(10, "-") & " PAGO EN EFECTIVO  " & String(10, "-")
                vsCadImpresion = vsCadImpresion & Space(16) & String(10, "-") & " PAGO EN EFECTIVO  " & String(10, "-") & lsSaltoLin
            End If
        End If
    End If
    
    vsCadImpresion = vsCadImpresion & "Capital :" & Space(13) & ImpreFormat(pnCapPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(14) & "Capital :" & Space(13) & ImpreFormat(pnCapPagado, 12) & lsSaltoLin
    
    If pnInteresDesagio > 0 Then
        vsCadImpresion = vsCadImpresion & "Interes Desagio (-)  :" & ImpreFormat(pnInteresDesagio, 12)
        vsCadImpresion = vsCadImpresion & Space(14) & "Interes Desagio (-)   :" & ImpreFormat(pnInteresDesagio, 12) & lsSaltoLin
    Else
        vsCadImpresion = vsCadImpresion & "Interes Compensatorio:" & ImpreFormat(pnIntCompPagado + pnIntCompVencPagado, 12)
        vsCadImpresion = vsCadImpresion & Space(14) & "Interes Compensatorio:" & ImpreFormat(pnIntCompPagado + pnIntCompVencPagado, 12) & lsSaltoLin
    End If
    
    vsCadImpresion = vsCadImpresion & "Mora :" & Space(16) & ImpreFormat(pnIntMorPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(14) & "Mora :" & Space(16) & ImpreFormat(pnIntMorPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Gastos :" & Space(14) & ImpreFormat(pnGastoPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(14) & "Gastos :" & Space(14) & ImpreFormat(pnGastoPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Int. Gracia:" & Space(10) & ImpreFormat(pnIntGraciaPagado, 12)
    vsCadImpresion = vsCadImpresion & Space(14) & "Int. Gracia:" & Space(10) & ImpreFormat(pnIntGraciaPagado, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Int. Susp y Reprog:" & Space(3) & ImpreFormat(pnIntSuspReprog, 12)
    vsCadImpresion = vsCadImpresion & Space(14) & "Int. Susp y Reprog:" & Space(3) & ImpreFormat(pnIntSuspReprog, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "I.T.F.            :" & Space(3) & ImpreFormat(pnITF, 12)
    vsCadImpresion = vsCadImpresion & Space(14) & "I.T.F.            :" & Space(3) & ImpreFormat(pnITF, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & String(40, "-")
    vsCadImpresion = vsCadImpresion & Space(12) & String(40, "-") & lsSaltoLin
    vsCadImpresion = vsCadImpresion & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntCompVencPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog + pnITF - pnInteresDesagio, 12)
    vsCadImpresion = vsCadImpresion & Space(14) & "Total Pagado :" & Space(8) & ImpreFormat(pnCapPagado + pnIntCompPagado + pnIntCompVencPagado + pnIntMorPagado + pnGastoPagado + pnIntGraciaPagado + pnIntSuspReprog + pnITF - pnInteresDesagio, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & lsNegritaOff
    vsCadImpresion = vsCadImpresion & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12)
    vsCadImpresion = vsCadImpresion & Space(14) & "Saldo Capital :" & Space(7) & ImpreFormat(pnSaldoCap, 12) & lsSaltoLin
    vsCadImpresion = vsCadImpresion & ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(3) & psCodUsu, 51, 0) & ImpreFormat("Prox. Fecha Pago : " & psProxFecPago & Space(3) & psCodUsu, 53, 0) & lsSaltoLin
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
    EspVerxPag = 56
    nPuntPag = 0
    lsTab = Space(1)
End Sub

Private Sub ImprimeCabeceraDocumento(ByRef psCadImp As String, ByVal psNomAge As String, _
    ByVal psFechaHora As String, ByVal psCodUsu, ByVal psTitulo As String, ByVal psNomCmac As String, _
    Optional psTab As String = "", Optional pbCondensado As Boolean = True, Optional psCodRepo As String = "", _
    Optional psSubTitulo As String = "") 'DAOR 20070314, Se aumentó subtitulo
    nPuntPag = 0
    psCadImp = psCadImp & lnSaltoLinDoc
    If Len(psCodRepo) > 0 Then
        psTitulo = psCodRepo & " - " & psTitulo
    End If
    If pbCondensado Then
        psCadImp = psCadImp & psTab & psNomCmac & Space(72) & "Fecha   : " & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 45, 0) & Space(51) & "Usuario : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
        psCadImp = psCadImp & psTab & Space((EspHorxPag - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & Chr$(10) 'DAOR 20070314
    Else
        psCadImp = psCadImp & psTab & psNomCmac & Space(45) & "Fecha :" & psFechaHora & Chr$(10)
        psCadImp = psCadImp & psTab & ImpreFormat(psNomAge, 40, 0) & Space(36) & "USUARIO : " & psCodUsu & Chr$(10)
        psCadImp = psCadImp & psTab & Space((120 - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
        psCadImp = psCadImp & psTab & Space((120 - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & Chr$(10) 'DAOR 20070314
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
            sCadImp = sCadImp & "ANALISTA : " & PstaNombre(RAnalista!cpersnombre)
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
    ByVal MatPagos As Variant, ByVal MatCuotasPend As Variant, ByVal MatDeudaVenc As Variant, ByVal MatDeudaAFecha As Variant, ByVal psProtesto As String, _
    ByVal psCargoAutom As String, ByVal psRefinanciado As String, ByVal psNomCmac As String, ByVal pMatGarantias As Variant, _
    ByVal pnCuotaCom As Integer, ByVal pnMiViv As Integer, ByVal pnCalendDin As Integer, ByVal psTasaComp As String, ByVal psTasaMor As String, ByVal MatRefinan As Variant) As String
    
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
    'CUOTAS PAGADAS
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "CUOTAS PAGADAS" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(157, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(9) & ImpreFormat("CUOTA", 10, 0) & ImpreFormat("FECHA PAG.", 17, 0) & ImpreFormat("TOTAL PAG.", 14, 0) & ImpreFormat("CAPITAL", 19, 0) & ImpreFormat("INTERES", 12, 0) & ImpreFormat("MORA", 8, 2) & ImpreFormat("GASTOS", 18, 0) & ImpreFormat("ITF", 10, 0) & ImpreFormat("ATRASO", 10, 0) & ImpreFormat("SALDO.K.", 10, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(157, "-") & lnSaltoLinDoc
    For I = 0 To UBound(MatPagos) - 1
        '                               Nro Cuota                           FECHA DE VENCIMIENTO                       TOTAL PAGADO                              CAPITAL                                   INTERES                                   MORA                                      GASTOS                                    ITF                                       DIAS ATRASO                               Saldo Kapìtal
        sCadImp = sCadImp & Space(10) & ImpreFormat(MatPagos(I, 1), 8, 0) & ImpreFormat(MatPagos(I, 0), 18, 0) & ImpreFormat(MatPagos(I, 2), 15, 0) & ImpreFormat(MatPagos(I, 3), 18, 0) & ImpreFormat(MatPagos(I, 4), 14, 0) & ImpreFormat(MatPagos(I, 5), 7, 0) & ImpreFormat(MatPagos(I, 6), 14, 2) & ImpreFormat(MatPagos(I, 7), 10, 2) & ImpreFormat(MatPagos(I, 8), 8, 2) & ImpreFormat(MatPagos(I, 9), 15, 2) & lnSaltoLinDoc
    Next I
    
    'CUOTAS PENDIENTES
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & lsNegritaOn & "CUOTAS PENDIENTES" & lsNegritaOff & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(9) & ImpreFormat("CUOTA", 10, 0) & ImpreFormat("FECHA", 17, 0) & ImpreFormat("CAPITAL", 14, 0) & ImpreFormat("INTERES", 19, 0) & ImpreFormat("MORA", 12, 0) & ImpreFormat("DIAS ATRASO", 18, 0) & ImpreFormat("ITF", 18, 0) & ImpreFormat("SALDO K.", 18, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
    For I = 0 To UBound(MatCuotasPend) - 1
        '                   Nro Cuota                                 FECHA DE VENCIMIENTO                      CAPITAL                                   INTERES                                   MORA                                      DIAS ATRASO                                          ITF                                        SALDO K
        sCadImp = sCadImp & Space(10) & ImpreFormat(MatCuotasPend(I, 0), 8, 0) & ImpreFormat(MatCuotasPend(I, 1), 18, 0) & ImpreFormat(MatCuotasPend(I, 2), 15, 0) & ImpreFormat(MatCuotasPend(I, 3), 18, 0) & ImpreFormat(MatCuotasPend(I, 4), 18, 0) & ImpreFormat(MatCuotasPend(I, 5), 12, 0) & ImpreFormat(MatCuotasPend(I, 6), 14, 2) & ImpreFormat(MatCuotasPend(I, 7), 10, 2) & lnSaltoLinDoc
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
    'sCadImp = sCadImp & Chr$(12)
    
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
    
    If IsArray(MatRefinan) Then
        If UBound(MatRefinan) > 0 Then
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & Space(5) & " REFINANCIACIONES " & lnSaltoLinDoc
            sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & Space(5) & Space(10) & ImpreFormat("Credito", 20) & ImpreFormat("Capital", 20) & ImpreFormat("Interes", 15) & ImpreFormat("Otros", 12) & lnSaltoLinDoc
            sCadImp = sCadImp & Space(5) & String(120, "-") & lnSaltoLinDoc
            For I = 0 To UBound(MatRefinan) - 1
                sCadImp = sCadImp & Space(15) & ImpreFormat(MatRefinan(I, 0), 20) & ImpreFormat(MatRefinan(I, 1), 20) & ImpreFormat(MatRefinan(I, 2), 15) & ImpreFormat(MatRefinan(I, 3), 12) & lnSaltoLinDoc
            Next I
            
        End If
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
Dim R2 As ADODB.Recordset
Dim sCadImp As String
Dim oCred As DCredito

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
        sCadImp = sCadImp & lsTab & "Cliente        : " & R!cPersCod & Space(5) & PstaNombre(R!cpersnombre) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Doc. Natural   : " & ImpreFormat(IIf(IsNull(R!DNI), "", R!DNI), 12, 0) & "Doc. Juridico    : " & ImpreFormat(IIf(IsNull(R!RUC), "", R!RUC), 12, 0) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Direccion      : " & R!cPersDireccDomicilio & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Nro Hijos      : " & Trim(Str(IIf(IsNull(R!nPersNatHijos), 0, R!nPersNatHijos))) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Edad           : " & Format(Year(CDate(psFecSis)) - Year(IIf(IsNull(R!dPersNacCreac), CDate(psFecSis), R!dPersNacCreac)), "#0") & " Años " & Space(20) & "Estado Civil : " & R!cEstadoCivil & lnSaltoLinDoc & lnSaltoLinDoc
       ' sCadImp = sCadImp & lsTab & lsNegritaOn & "FUENTE DE INGRESO" & Space(30) & IIf(IsNull(R!cFteInd), "DEPENDIENTE", "INDEPENDIENTE") & lsNegritaOff & lnSaltoLinDoc & lnSaltoLinDoc
         sCadImp = sCadImp & lsTab & "FUENTE DE INGRESO" & Space(30) & IIf(CInt((R!nTipoFuente)) = 1, "DEPENDIENTE", "INDEPENDIENTE") & lnSaltoLinDoc & lnSaltoLinDoc
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
        
        Set oCred = New DCredito
        Set R2 = oCred.RecuperaCreditosRefinanciados(psCtaCod)
        If R2.RecordCount > 0 Then
            sCadImp = sCadImp & lsTab & "Creditos a Refinanciar :" & lnSaltoLinDoc & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("Credito", 20) & ImpreFormat("Capital", 10) & ImpreFormat("Interes", 20) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
            Do While Not R2.EOF
                sCadImp = sCadImp & lsTab & ImpreFormat(R2!cCtaCodref, 18) & ImpreFormat(R2!nCapitalRef, 8) & ImpreFormat(R2!nInteresRef, 8) & lnSaltoLinDoc
                R2.MoveNext
            Loop
            sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
        End If
        R2.Close
        Set oCred = Nothing
                
        
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
Dim odCredDoc As DCredDoc
Dim oDCred As DCredito
Dim oDGarantia As DGarantia
Dim R As ADODB.Recordset
Dim RGarantCred As ADODB.Recordset
Dim RRelaCred As ADODB.Recordset
Dim REstadosCred As ADODB.Recordset
Dim sCadImp As String

Dim sEdadTiempoFunc As String

    Set odCredDoc = New DCredDoc
    Set R = odCredDoc.RecuperaDatosResumenComite(psCtaCod)
    Set odCredDoc = Nothing
    Set oDCred = New DCredito
    Set RRelaCred = oDCred.RecuperaRelacPers(psCtaCod)
    Set REstadosCred = oDCred.RecuperaEstadosdelCredito(psCtaCod, True, True, False)
    Set oDCred = Nothing
    Set oDGarantia = New DGarantia
    Set RGarantCred = oDGarantia.RecuperaGarantiaCredito(psCtaCod)
    Set oDGarantia = Nothing
        
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "RESUMEN DE COMITE", psNomCmac, lsTab, False)
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & lsNegritaOn & "Credito : " & psCtaCod & ImpreFormat(IIf(Mid(psCtaCod, 9, 1) = "1", "SOLES", "DOLARES"), 5, 5) & Space(25) & R!cTipoCred & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "DATOS DEL TITULAR" & Space(30) & R!cPersoneria & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Cliente        : " & R!cPersCod & Space(2) & PstaNombre(R!cpersnombre) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Doc. Natural   : " & R!DNI & Space(30) & "Doc. Juridico : " & R!RUC & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "Direccion      : " & R!cPersDireccDomicilio & lnSaltoLinDoc
    
    ' CMACICA_CSTS - 18/11/2003 ------------------------------------------------------------------------------------
    
    If R!nPersPersoneria <> 1 Then
       sEdadTiempoFunc = "Tiempo de Func.: "
    Else
       sEdadTiempoFunc = "Edad           : "
    End If
    
    If IsNull(R!dPersNacCreac) Then
        sCadImp = sCadImp & lsTab & sEdadTiempoFunc
    Else
        sCadImp = sCadImp & lsTab & sEdadTiempoFunc & Format(Year(CDate(psFecSis)) - Year(R!dPersNacCreac), "#0") & " Años "
    End If
    '----------------------------------------------------------------------------------------------------------------
    
    
    ' CMACICA_CSTS - 18/11/2003 ------------------------------------------------------------------------------------
    If R!nPersPersoneria <> 1 Then
       sCadImp = sCadImp & Space(45) & lnSaltoLinDoc
    Else
       sCadImp = sCadImp & Space(30) & "Estado Civil : " & R!cEstadoCivil & lnSaltoLinDoc
    End If
    '----------------------------------------------------------------------------------------------------------------
    
    ' CMACICA_CSTS - 18/11/2003 ------------------------------------------------------------------------------------
    If R!nPersPersoneria <> 1 Then
       sCadImp = sCadImp & lnSaltoLinDoc
    Else
       sCadImp = sCadImp & lsTab & "Nro Hijos : " & Trim(Str(IIf(IsNull(R!nPersNatHijos), 0, R!nPersNatHijos))) & lnSaltoLinDoc
    End If
    '----------------------------------------------------------------------------------------------------------------
    
    sCadImp = sCadImp & lsNegritaOn
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & Space(35) & "CLIENTES RELACIONADOS" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat(" CODIGO ", 15, 0) & ImpreFormat("NOMBRE", 40, 0) & ImpreFormat("RELACION", 20, 0) & ImpreFormat("DOC. IDENT", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsNegritaOff
    Do While Not RRelaCred.EOF
        If RRelaCred!nConsValor <> gColRelPersTitular Then
            sCadImp = sCadImp & lsTab & ImpreFormat(RRelaCred!cPersCod, 15, 0) & ImpreFormat(RRelaCred!cpersnombre, 40, 0) & ImpreFormat(RRelaCred!cConsDescripcion, 20, 0) & ImpreFormat(RRelaCred!DNI, 15, 0) & lnSaltoLinDoc
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
Dim sPersNombreUser As String

    Set oCredDoc = New DCredDoc
    Set R = oCredDoc.RecuperaDatosDocDesembolso(psCtaCod)
    sPersNombreUser = oCredDoc.ObtieneNombrePersonaUser(psCodUsu)

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
    sCadImp = sCadImp & lsTab & "Fecha de Desembolso : " & ImpreFormat(Format(R!dFecDesemb, "dd/mm/yyyy"), 20, 0) & "Nombre            : " & PstaNombre(R!cpersnombre) & lnSaltoLinDoc
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
    sCadImp = sCadImp & lsTab & ImpreFormat("F. PACTADA", 15, 0) & ImpreFormat("FECHA", 15, 0) & ImpreFormat("CUOTA", 15, 0) & ImpreFormat("MONTO", 15, 0) & ImpreFormat("GASTOS", 15, 0) & ImpreFormat("ESTADO", 15, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc
    Do While Not RDesemb.EOF
        sCadImp = sCadImp & lsTab & ImpreFormat(Format(RDesemb!dVenc, "dd/mm/yyyy"), 15, 0) & ImpreFormat(Format(RDesemb!dPago, "dd/mm/yyyy"), 15, 0) & ImpreFormat(Trim(Str(RDesemb!nCuota)), 5, 0) & ImpreFormat(RDesemb!nCapital, 15, 2) & ImpreFormat(RDesemb!NGasto, 10, 2) & Space(10) & IIf(RDesemb!ncoloccalendestado = gColocCalendEstadoPendiente, "PENDIENTE", "DESEMBOLSADO") & lnSaltoLinDoc
        RDesemb.MoveNext
    Loop
    RDesemb.Close
    Set RDesemb = Nothing
    sCadImp = sCadImp & lsTab & String(90, "-") & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "________________________________   ______________________________    _________________________" & lnSaltoLinDoc
    'sCadImp = sCadImp & lsTab & PstaNombre(R!cPersNombre) & Space(5) & PstaNombre(R!cAnalista) & Space(15) & "CAJA" & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & PstaNombre(R!cpersnombre) & Space(5) & PstaNombre(sPersNombreUser) & Space(19) & lnSaltoLinDoc
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
Dim bEstadoConvenio As Boolean

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
    sCadImp = sCadImp & lsTab & "Credito       : " & psCtaCod & Space(10) & "Cliente : " & PstaNombre(R!cpersnombre) & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & lsTab & "Analista      : " & PstaNombre(R!cAnalista) & Chr$(10)
    sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & ImpreFormat(R!cTipoCuota, 30, 0) & "Cuota : " & Format(R!nMontoCuota, "#0.00") & Chr$(10)
    sCadImp = sCadImp & lsTab & "Interes       : " & ImpreFormat(Format(R!nTasaInteres, "#0.00"), 30, 0) & "Plazo : " & Trim(Str(R!nPlazo)) & Chr$(10)
    sCadImp = sCadImp & lsTab & "Moneda        : " & ImpreFormat(R!cMoneda, 30, 0) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & Chr$(10)
    sCadImp = sCadImp & lsTab & "Monto         : " & ImpreFormat(Format(R!nMonto, "#0.00"), 30, 0) & "Gracia : " & Trim(Str(R!nPeriodoGracia)) & Chr$(10)
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
            sCadImp = sCadImp & lsTab & ImpreFormat(Format(RDesPar!dVenc, "dd/mm/yyyy"), 14, 0) & ImpreFormat(RDesPar!nCapital, 10, 2, True) & ImpreFormat(IIf(RDesPar!ncoloccalendestado = 1, "DESEMBOLSADO", "PENDIENTE"), 12, 2, True) & Chr$(10)
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
        If I <> 0 And I Mod 44 = 0 Then
            sCadImp = sCadImp & Chr$(12)
            Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", ImprimePlandePagos, lsTab, False)
            sCadImp = sCadImp & lsTab & "DUPLICADO" & Chr$(10)
            sCadImp = sCadImp & lsTab & "Credito       : " & psCtaCod & Space(10) & "Cliente : " & PstaNombre(R!cpersnombre) & Chr$(10) & Chr$(10)
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
    
    'function por credicom para verificar si tiene o no convenion con los gastos
    Set oCredDoc = New DCredDoc
    bEstadoConvenio = oCredDoc.VerificaConvenioEmp(psCtaCod)
    Set oCredDoc = Nothing
    
    If bEstadoConvenio = True Then
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & lsTab & " Se debe indicar que los siguientes costos por comision de servicios (COM) y Aportes  de" & Chr$(10)
        sCadImp = sCadImp & lsTab & " los socios (A. Socios) no forman parte de los interes ni gastos que originan el crédito" & Chr$(10)
        sCadImp = sCadImp & lsTab & " del presente cronograma en tal sentido la CAJA no tiene ninguna  responsabilidad por el" & Chr$(10)
        sCadImp = sCadImp & lsTab & " uso o destino  de dichos conceptos." & Chr$(10)
    End If
    
    sCadImp = sCadImp & Chr$(12)
    ImprimePlandePagos = sCadImp
End Function

Public Function ImprimeCalendarioDesemb(ByVal psCtaCod As String, ByVal psNomAge As String, _
    ByVal psFecSis As String, ByVal psCodUsu As String, ByVal pnMontoPrestamo As Double, _
    ByVal pbMiViv As Boolean, Optional ByVal pbCalendParalelo As Boolean = False, Optional ByVal psNomCmac As String = "", _
    Optional ByVal pbComodin As Integer = 0, Optional ByVal pbCalendDin As Integer = 0) As String

Dim R As ADODB.Recordset
Dim oCredDoc As DCredDoc
Dim oCalend As Dcalendario
Dim oNCred As NCredito
Dim MatCalend As Variant
Dim sCadImp As String
Dim nCapital As Double
Dim RDesPar As ADODB.Recordset

    Set oCredDoc = New DCredDoc
    Set R = oCredDoc.RecuperaDatosDocPlanPagos(psCtaCod)
    Set oCredDoc = Nothing
    
   
    Set oCalend = New Dcalendario
    Set RDesPar = oCalend.RecuperaCalendarioDesemb(psCtaCod)
    Set oCalend = Nothing
    
    Set oNCred = New NCredito
    MatCalend = oNCred.RecuperaMatrizCalendarioInicial(psCtaCod, pnMontoPrestamo, pbCalendParalelo, IIf(pbCalendParalelo, True, False))
    Set oNCred = Nothing
    
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUsu, "PLAN DE PAGOS", psNomCmac, lsTab, False)
    sCadImp = sCadImp & lsTab & "DUPLICADO" & Chr$(10)
    sCadImp = sCadImp & lsTab & "Credito       : " & psCtaCod & Space(10) & "Cliente : " & PstaNombre(R!cpersnombre) & Chr$(10) & Chr$(10)
    sCadImp = sCadImp & lsTab & "Analista      : " & PstaNombre(R!cAnalista) & Chr$(10)
    sCadImp = sCadImp & lsTab & "Tipo de Cuota : " & ImpreFormat(R!cTipoCuota, 30, 0) & "Cuota : " & Format(R!nMontoCuota, "#0.00") & Chr$(10)
    sCadImp = sCadImp & lsTab & "Interes       : " & ImpreFormat(Format(R!nTasaInteres, "#0.00"), 30, 0) & "Plazo : " & Trim(Str(R!nPlazo)) & Chr$(10)
    sCadImp = sCadImp & lsTab & "Moneda        : " & ImpreFormat(R!cMoneda, 30, 0) & "Vigencia : " & Format(R!dFecVig, "ddd, d mmm yyyy") & Chr$(10)
    sCadImp = sCadImp & lsTab & "Monto         : " & Format(R!nMonto, "#0.00") & Chr$(10)

    sCadImp = sCadImp & lsTab & "Tipo Calend.  : " & ImpreFormat("Normal", 30, 0)
    
    sCadImp = sCadImp & lsTab & "Calend. Din.  :" & ImpreFormat("NO", 30, 0)

    sCadImp = sCadImp & Chr$(10)
    
    sCadImp = sCadImp & Chr$(10)
    sCadImp = sCadImp & Chr$(10)
    
    If RDesPar.RecordCount >= 2 Then
        sCadImp = sCadImp & lsTab & "DESEMBOLSOS PARCIALES" & Chr$(10) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(80, "-") & Chr$(10)
        sCadImp = sCadImp & lsTab & ImpreFormat("FECHA", 16, 0) & ImpreFormat("DESEMBOLSO", 10, 0) & ImpreFormat("ESTADO", 10, 2) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(80, "-") & Chr$(10)
        Do While Not RDesPar.EOF
            sCadImp = sCadImp & lsTab & ImpreFormat(Format(RDesPar!dVenc, "dd/mm/yyyy"), 14, 0) & ImpreFormat(RDesPar!nCapital, 10, 2, True) & ImpreFormat(IIf(RDesPar!ncoloccalendestado = 1, "DESEMBOLSADO", "PENDIENTE"), 12, 2, True) & Chr$(10)
            RDesPar.MoveNext
        Loop
    End If
    RDesPar.Close
    
    sCadImp = sCadImp & Chr$(10)
    
   
    sCadImp = sCadImp & Chr$(27) & Chr$(84) & Chr$(1)
    sCadImp = sCadImp & Chr$(27) + Chr$(107) + Chr$(1)
    
    sCadImp = sCadImp & Chr$(12)
    ImprimeCalendarioDesemb = sCadImp
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
Dim nAcumITFPag As Double

'Agregado por LMMD
Dim nProxCuota As String
Dim nProFecVenc As String
Dim rs As ADODB.Recordset

    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecha, psCodUsu, "Reporte de Pagos de Prestamos Personales", psNomCmac, lsTab, True)
    sCadImp = sCadImp & lsTab & "INSTITUCION : " & psInstitucion & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "MONEDA : " & psMoneda & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(165, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 25, 0) & ImpreFormat("TITULAR", 30, 0) & ImpreFormat("CUOTA QUE PAGO", 20, 0) & ImpreFormat("MONTO", 15, 0) & ImpreFormat("CAPITAL", 15, 0) & ImpreFormat("INTERES", 12, 0) & ImpreFormat("MORA", 8, 0) & ImpreFormat("GASTOS", 8, 0) & ImpreFormat("I.T.F.", 6, 2) & Space(2) & ImpreFormat("P.Cuota", 8, 0) & ImpreFormat("F.PROX", 12) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(165, "-") & lnSaltoLinDoc
    
    Set oNCredito = New NCredito
    nAcumTot = 0
    nAcumCapPag = 0
    nAcumIntComPag = 0
    nAcumIntMorPag = 0
    nAcumGastosPag = 0
    For I = 0 To UBound(MatPagos) - 1
        nMontoPag = CDbl(MatPagos(I, 1)) + CDbl(MatPagos(I, 9))
        sCadImp = sCadImp & lsTab & ImpreFormat(MatPagos(I, 0), 21) & ImpreFormat(MatPagos(I, 3), 30) & ImpreFormat(IIf(oNCredito.MatrizCuotasPagadas(MatCalendDistrib(I)) = 0, ObtenerCuotaPagada(MatPagos(I, 0)), oNCredito.MatrizCuotasPagadas(MatCalendDistrib(I))), 10) _
        & ImpreFormat(nMontoPag, 12) & ImpreFormat(oNCredito.MatrizCapitalPagado(MatCalendDistrib(I)), 12) & ImpreFormat(oNCredito.MatrizIntCompPagado(MatCalendDistrib(I)) + oNCredito.MatrizIntCompVencPagado(MatCalendDistrib(I)) + oNCredito.MatrizIntGraciaPagado(MatCalendDistrib(I)) + oNCredito.MatrizIntReprogPag(MatCalendDistrib(I)) + oNCredito.MatrizIntSuspensoPag(MatCalendDistrib(I)), 10) & ImpreFormat(oNCredito.MatrizIntMorPagado(MatCalendDistrib(I)), 8) & ImpreFormat(oNCredito.MatrizGastoPag(MatCalendDistrib(I)), 6) & ImpreFormat(CDbl(MatPagos(I, 9)), 6, 2)
        
        nProxCuota = ""
        nProFecVenc = ""
         Set rs = ObtenerProxCuota(MatPagos(I, 0))
         If Not rs.EOF And Not rs.BOF Then
            nProxCuota = rs!nCuota
            nProFecVenc = Format(rs!dVenc, "dd/mm/yyyy")
         End If
         Set rs = Nothing
         sCadImp = sCadImp & Space(2) & ImpreFormat(nProxCuota, 3, 0) & ImpreFormat(nProFecVenc, 10) & lnSaltoLinDoc
         
         nAcumTot = nAcumTot + nMontoPag
         nAcumCapPag = nAcumCapPag + oNCredito.MatrizCapitalPagado(MatCalendDistrib(I))
         nAcumIntComPag = nAcumIntComPag + oNCredito.MatrizIntCompPagado(MatCalendDistrib(I)) + oNCredito.MatrizIntCompVencPagado(MatCalendDistrib(I)) + oNCredito.MatrizIntGraciaPagado(MatCalendDistrib(I)) + oNCredito.MatrizIntReprogPag(MatCalendDistrib(I)) + oNCredito.MatrizIntSuspensoPag(MatCalendDistrib(I))
         nAcumIntMorPag = nAcumIntMorPag + oNCredito.MatrizIntMorPagado(MatCalendDistrib(I))
         nAcumGastosPag = nAcumGastosPag + oNCredito.MatrizGastoPag(MatCalendDistrib(I))
         nAcumITFPag = nAcumITFPag + CDbl(MatPagos(I, 9))
    Next I
    Set oNCredito = Nothing
    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " PAGO TOTAL     : " & ImpreFormat(nAcumTot, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " CAPITAL PAGADO : " & ImpreFormat(nAcumCapPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " INTERES PAGADO : " & ImpreFormat(nAcumIntComPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " MORA PAGADA    : " & ImpreFormat(nAcumIntMorPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " GASTO PAGADO   : " & ImpreFormat(nAcumGastosPag, 10, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " ITF PAGADO     : " & ImpreFormat(nAcumITFPag, 10, 2) & lnSaltoLinDoc
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
Dim saux As Integer
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

Dim lsmensaje As String * 39

''  quitar este comentario

        Dim clsGen As DGeneral
        Set clsGen = New DGeneral
 

ETIQ:

On Error GoTo Error

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
saux = 5

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
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Chr$(10)
        End If
        
        If psNomAgeRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & Chr$(10)
        End If
        
        If psComCmac = "XXX" Then
            If psLin3 = "XXX" Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff 'Desactiva Negrita
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & lsNegritaOff   'Desactiva Negrita
                lnNumLinCmac = 1
            End If
            lnNumLinCmac = 0
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & lsNegritaOff & Chr$(10)   'Desactiva Negrita
            lnNumLinCmac = 1
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & Chr$(10)
        
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & Chr$(10)
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
            
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & Chr$(10)
            
            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + saux) & Chr$(10)
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
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & Chr$(10)
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
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(sCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & Space(sSep) & Chr$(10)
        End If
        'Dim clsGen As DGeneral
        'Set clsGen = New DGeneral
        lsmensaje = clsGen.GetMensajeBoletas(psCodCta)
        Set clsGen = Nothing
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & lsNegritaOn & ImpreCarEsp(lsmensaje) & Space(-1 + sSep) & ImpreCarEsp(lsmensaje) & lsNegritaOff
        
        lnNumLinCmac = lnNumLinCmac + 1
        
        For saux = 1 To (lnTope - lnNumLinCmac)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Next saux

Else   '' Caja Trujillo
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOn  'Activa Negrita
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "CM - AHORRO" & Space(19 + sSep + saux) & "CMACT - AHORRO" & Chr$(10)
        
        If Mid(psCodCta, 9, 1) = 1 Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Space(sSep) & Trim(sNomAge) & "-SOLES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt & Space(sSep) & Trim(sNomAge) & "-DOLARES" & Space(saux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt & Chr$(10)
        End If
        
        If psNomAgeRem = "" Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Ag.Rem: " & Trim(psNomAgeRem) & Space(saux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))) & "Ag.Rem: " & Trim(psNomAgeRem) & Chr$(10)
        End If
        
        If psComCmac = "XXX" Then
            If psLin3 = "XXX" Then
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & lsNegritaOff 'Desactiva Negrita
            Else
                ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & psLin3 & Space(saux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff  'Desactiva Negrita
                lnNumLinCmac = 1
            End If
            lnNumLinCmac = 0
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "NroDocCmac:" & psComCmac & Space(saux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff & Chr$(10)   'Desactiva Negrita
            lnNumLinCmac = 1
        End If
        
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Space(saux + sSep - 6) & "Fecha:" & sFecha & Space(10) & "Hora:" & sHora & Chr$(10)
        
        lnCliAux = InStr(1, psCliente, "*", vbTextCompare)
        
        If lnCliAux = 0 Then
            If saux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + saux)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(psCliente) & Space(saux + sMax + sSep - Len(psCliente)) & ImpreCarEsp(psCliente) & Chr$(10)
        Else
            lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
            lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
            
            If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + saux)
            If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + saux)
            
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(lsCliAux1) & Space(saux + sMax + sSep - Len(lsCliAux1)) & ImpreCarEsp(lsCliAux1) & Chr$(10)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(lsCliAux2) & Space(saux + sMax + sSep - Len(lsCliAux2)) & ImpreCarEsp(lsCliAux2) & Chr$(10)
            
            lnCliAux = 1
        End If
        
        If pbSaldoInt Or pbCuenta Then
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Cuenta:" & psCodCta & Space(8 + sSep + saux) & "Cuenta:" & psCodCta & Chr$(10)
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
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(Trim(psTexto)) & Space(saux + sSep + sMax - Len(Trim(psTexto))) & ImpreCarEsp(Trim(psTexto)) & Chr$(10)
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
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & ImpreCarEsp(sCodUser) & Space(29 + sSep + saux) & ImpreCarEsp(sCodUser) & Chr$(10)
        Else
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(sCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & Space(sSep) & "Loc/Rem" & Space(sMax + saux - Len("Loc/Rem") - 1 - 8) & ImpreCarEsp(sCodUser) & "/" & ImpreCarEsp(psCodUsuRem) & Chr$(10)
        End If
'        Dim clsGen As DGeneral
'        Set clsGen = New DGeneral
        lsmensaje = clsGen.GetMensajeBoletas(psCodCta)
        Set clsGen = Nothing
        ImprimeBoletaAhorro = ImprimeBoletaAhorro & Space(sIni) & lsNegritaOn & ImpreCarEsp(lsmensaje) & Space(-1 + sSep) & ImpreCarEsp(lsmensaje) & lsNegritaOff
        
        lnNumLinCmac = lnNumLinCmac + 1
        
        For saux = 1 To (lnTope - lnNumLinCmac)
            ImprimeBoletaAhorro = ImprimeBoletaAhorro & "" & Chr$(10)
        Next saux
        
End If
'''- Fin de Liquidaciones
If sLpt <> "" Then
    Print #nFicSal, ImprimeBoletaAhorro
    Print #nFicSal, Chr$(12)
    Close nFicSal
End If

Exit Function

Error:
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
    Optional ByVal psNomCmac As String = "", Optional ByVal psNomCmacAbrev As String = "", Optional ByVal psCodCMAC As String, Optional pnTIF As Double = 0) As Variant
    
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
Dim odCredDoc As DCredDoc
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
    
    Set odCredDoc = New DCredDoc
    Set R = odCredDoc.RecuperaDatosBoletaDesembolso(pnMovNro, psCtaCod, psNomAge)
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
            R!FechaTran, R!HoraTran, R!cpersnombre, Trim(Str(nNumTran)), psCtaCod, "", psNomCmacAbrev, psCodCMAC)
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
    
    'sCadTmp = sCadTmp & "Monto del Desembolso :" & Space(2) & ImpreFormat(nCapital + pnTIF, 12)
    sCadTmp = sCadTmp & "Monto del Desembolso :" & Space(2) & ImpreFormat(nCapital, 12)
    'sCadTmp = sCadTmp & Space(14) & "Monto del Desembolso :" & Space(2) & ImpreFormat(nCapital + pnTIF, 12) & lsSaltoLin
    sCadTmp = sCadTmp & Space(14) & "Monto del Desembolso :" & Space(2) & ImpreFormat(nCapital, 12) & lsSaltoLin
    sCadTmp = sCadTmp & "Gastos:" & ImpreFormat(Abs(NGasto), 29)
    sCadTmp = sCadTmp & Space(14) & "Gastos :" & ImpreFormat(Abs(NGasto), 28) & lsSaltoLin
    
    sCadTmp = sCadTmp & "ITF   :" & String(26, " ") & ImpreFormat(Format(Abs(pnTIF), "#0.00") & "(-)", 15)
    sCadTmp = sCadTmp & Space(2) & "ITF    :" & String(26, " ") & ImpreFormat(Format(Abs(pnTIF), "#.00") & "(-)", 28) & lsSaltoLin
    
    
    'ITF
    'sCadTmp = sCadTmp & lsSaltoLin
    
    
    sCadTmp = sCadTmp & lsSaltoLin
    sCadTmp = sCadTmp & String(45, "-")
    sCadTmp = sCadTmp & Space(7) & String(44, "-") & lsSaltoLin
    'sCadTmp = sCadTmp & "Total a Desembolsar :" & Space(8) & ImpreFormat(nCapital + NGasto, 12)
    'sCadTmp = sCadTmp & Space(15) & "Total a Desembolsar  :" & Space(8) & ImpreFormat(nCapital + NGasto, 12) & lsSaltoLin
    sCadTmp = sCadTmp & "Total a Entregar :" & Space(8) & ImpreFormat(nCapital + NGasto - pnTIF, 10)
    sCadTmp = sCadTmp & Space(15) & "Total a Entregar  :" & Space(8) & ImpreFormat(nCapital + NGasto - pnTIF, 12) & lsSaltoLin
    sCadTmp = sCadTmp & lsNegritaOff
    
     sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & ImpreFormat(nCapital, 12)
    'sCadTmp = sCadTmp & "Saldo Capital :" & Space(7) & ImpreFormat(nCapital + pnTIF, 12)
    'sCadTmp = sCadTmp & Space(16) & "Saldo Capital :" & Space(7) & ImpreFormat(nCapital + pnTIF, 13) & lsSaltoLin
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
    Set R = odCredDoc.RecuperaDatosBoletaGastosDesembolso(pnMovNro, psCtaCod, psNomAge)
    If Not R.BOF And Not R.EOF Then
        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
                R!FechaTran, R!HoraTran, R!cpersnombre, "", psCtaCod, "", psNomCmacAbrev, psCodCMAC)
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
        sCadTmp = sCadTmp & ImpreFormat(Mid(R!CDescripcion, 1, 25), 25) & ImpreFormat(Abs(R!nMonto), 12) & lsSaltoLin
     Else
        sCadTmp = sCadTmp & ImpreFormat(Mid(R!CDescripcion, 1, 25), 25) & ImpreFormat(Abs(R!nMonto), 12)
        sCadTmp = sCadTmp & Space(11) & ImpreFormat(Mid(R!CDescripcion, 1, 25), 25) & ImpreFormat(Abs(R!nMonto), 12) & lsSaltoLin
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
    Set R = odCredDoc.RecuperaDatosBoletaCancelCredDesembolso(pnMovNro, psCtaCod, psNomAge)
    ReDim MatTempDesemCancel(R.RecordCount, 2)
    Do While Not R.EOF
    sCadTmp = ""
    vsCadImpresion = ""
    If psCodCMAC = "102" Then '' Liquidacion lima
        Call ImprimeCabeceraBoletaPago(psNomAge, R!sMoneda, "", _
                R!FechaTran, R!HoraTran, R!cpersnombre, Trim(Str(R!nNumTran)), R!cCtaCod, "", psNomCmacAbrev, psCodCMAC)
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
                R!FechaTran, R!HoraTran, R!cpersnombre, Trim(Str(R!nNumTran)), R!cCtaCod, "", psNomCmacAbrev, psCodCMAC)
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
        
        sCadTmp = sCadTmp & "ITF     :" & Space(13) & ImpreFormat(IIf(IsNull(R!nitf), 0, R!nitf), 12)
        sCadTmp = sCadTmp & Space(16) & "ITF     :" & Space(13) & ImpreFormat(IIf(IsNull(R!nitf), 0, R!nitf), 12) & lsSaltoLin
        
        'sCadTmp = sCadTmp & lsSaltoLin
        sCadTmp = sCadTmp & lsSaltoLin
        sCadTmp = sCadTmp & String(40, "-")
        sCadTmp = sCadTmp & Space(14) & String(40, "-") & lsSaltoLin
        sCadTmp = sCadTmp & "Total a Pagar :" & Space(8) & ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital) + IIf(IsNull(R!nGastos), 0, R!nGastos) + IIf(IsNull(R!nInteres), 0, R!nInteres) + IIf(IsNull(R!nitf), 0, R!nitf), 12)
        sCadTmp = sCadTmp & Space(16) & "Total a Pagar :" & Space(8) & ImpreFormat(IIf(IsNull(R!nCapital), 0, R!nCapital) + IIf(IsNull(R!nInteres), 0, R!nInteres) + IIf(IsNull(R!nGastos), 0, R!nGastos) + IIf(IsNull(R!nitf), 0, R!nitf), 12) & lsSaltoLin
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
    MatTempPlan(0, 1) = ""
    For I = 1 To 3
        MatTempPlan(0, 1) = MatTempPlan(0, 1) & ImprimePlandePagos(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu, nMontoDesemb, nCalMiVivi)
        If nCalMiVivi = True Then
            MatTempPlan(0, 1) = MatTempPlan(0, 1) & ImprimePlandePagos(psCtaCod, psNomAge, Format(pdFecSis, "dd/mm/yyyy"), psCodUsu, nMontoDesemb, nCalMiVivi, True)
        End If
    Next I
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
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "", Optional ByVal pnFRA As Integer = 0) As String
    
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset
Dim vsProducto As String
Dim vsFinanc As String
Dim vsOperacion As String
Dim nTotal, nTotalCap, nTotalInt, nTotalMor, nTotalGastos As Double

    Set oDCred = New DCredDoc
    Set R = oDCred.ReportePagoDeCreditos(pMatAgencias, pdFecIni, pdFecFin, pnMoneda, pnFRA)
    Set oDCred = Nothing
    sCadImp = ""
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "INGRESOS POR PAGO DE CREDITOS DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
    
    If R.RecordCount > 0 Then
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("IMPORTE", 13) & ImpreFormat("CAPITAL", 13) & ImpreFormat("INTERES", 16) & ImpreFormat("MORA", 12) & ImpreFormat("GASTOS", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    End If
    
    nTotal = 0
    nTotalCap = 0
    nTotalInt = 0
    nTotalMor = 0
    nTotalGastos = 0
    
    Do While Not R.EOF
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!cpersnombre, 20) & _
                                    ImpreFormat(R!Capital + R!Interes + R!Mora + R!Gastos, 20) & _
                                    ImpreFormat(R!Capital, 12) & ImpreFormat(R!Interes, 12) & ImpreFormat(R!Mora, 12) & ImpreFormat(R!Gastos, 12) & lnSaltoLinDoc
                        
    
        nTotal = nTotal + R!Capital + R!Interes + R!Mora + R!Gastos
        nTotalCap = nTotalCap + R!Capital
        nTotalInt = nTotalInt + R!Interes
        nTotalMor = nTotalMor + R!Mora
        nTotalGastos = nTotalGastos + R!Gastos
                        
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(53) & ImpreFormat(nTotal, 12) & ImpreFormat(nTotalCap, 12) & ImpreFormat(nTotalInt, 12) & ImpreFormat(nTotalMor, 12) & ImpreFormat(nTotalGastos, 12) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(130, "-") & lnSaltoLinDoc
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
        'sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
        '---------
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!copecod) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 19) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("PRESTAMO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 8) & ImpreFormat("ANALISTA", 10) & ImpreFormat("INTERES", 12) & ImpreFormat("USER", 4) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE OPERACION
        If Trim(R!NCredito) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL OPERACION ", 42) & _
                                    ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE FUENTE DE FINANCIAMIENTO
            If Trim(R!copecod) = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                'sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FUENTE FINANCIAMIENTO ", 42) & _
                                    ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL ANALISTA", 42) & _
                                    ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & lnSaltoLinDoc
                '---------------
                sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If R!sFuenteFinan = "ZTOTAL" Then
                    sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 42) & _
                                    ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    If R!sProducto = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL ", 42) & _
                                    ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'TOTAL AGENCIA
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & _
                                        ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & ImpreFormat(R!nTasa, 18, 4) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
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
                            'sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                            '------
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!copecod) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 19) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 8) & ImpreFormat("ANALISTA", 10) & ImpreFormat("INTERES", 12) & ImpreFormat("USER", 4) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        'sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                        '-------
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!copecod) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 19) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 8) & ImpreFormat("ANALISTA", 10) & ImpreFormat("INTERES", 12) & ImpreFormat("USER", 4) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                    End If
                Else
                    'sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                    '----------
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!copecod) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 19) & ImpreFormat("CLIENTE", 34) & ImpreFormat("APROBADO", 13) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 8) & ImpreFormat("ANALISTA", 10) & ImpreFormat("INTERES", 12) & ImpreFormat("USER", 4) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "OPERACION : " & Trim(R!copecod) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 19) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("APROBADO", 13) & ImpreFormat("INT APROB", 16) & ImpreFormat("GASTO", 8) & ImpreFormat("ANALISTA", 10) & ImpreFormat("INTERES", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(150, "-") & lnSaltoLinDoc
            End If
        End If
                
        sCadImp = sCadImp & lsTab & ImpreFormat(R!NCredito, 20) & ImpreFormat(R!sTitular, 20) & _
                    ImpreFormat(R!nDesembolso, 22) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nIntApr, 10) & ImpreFormat(R!Gasto, 14) & ImpreFormat(R!sAnalista, 6, 6) & ImpreFormat(R!nTasa, 6, 4) & ImpreFormat(R!cUserDes, 6) & lnSaltoLinDoc
        
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
        If Trim(R!sFuenteFinan) = "TOTAL" Then
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
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!sFuenteFinan, 20) & ImpreFormat(R!nCasos, 15, 0) & _
                                    ImpreFormat(R!nMontoApr, 12) & ImpreFormat(R!nDesembolsado, 12) & ImpreFormat(R!nSaldoCapital, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeSaldoCarteraVigente = sCadImp
End Function


Public Function ImprimeCreditosCancelados(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
    ByVal pdFecFin As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "", _
    Optional ByVal pMatAnalistas As Variant)
    
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset
Dim I As Integer
Dim sAnalista As String

    For I = 0 To UBound(pMatAnalistas) - 1
        If I = 0 Then
         sAnalista = "'" & pMatAnalistas(I) & "'"
        Else
         sAnalista = sAnalista & ",'" & pMatAnalistas(I) & "'"
        End If
    Next I

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosCanceladosTodos(pMatAgencias, pdFecIni, pdFecFin, pnMoneda, sAnalista)
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
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CREDITO ANTIGUO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
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
            If R!sFuenteFinan = "TOTAL" Then
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
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CREDITO ANTIGUO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CREDITO ANTIGUO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CREDITO ANTIGUO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CREDITO ANTIGUO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DESEMBOLSO", 13) & ImpreFormat("Nro CUOTAS", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("VIGENCIA", 15) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!NCredito, 20) & ImpreFormat(R!CuentaAntigua, 20) & ImpreFormat(R!sTitular, 20) & _
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
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
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
            If R!sFuenteFinan = "TOTAL" Then
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
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                            sCadImp = sCadImp & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("SALDO", 13) & ImpreFormat("CAPITAL", 12) & ImpreFormat("INTERES", 12) & ImpreFormat("MORA", 8) & ImpreFormat("INT. PERD", 12) & ImpreFormat("LINEA C.", 8) & ImpreFormat("ANALISTA", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & lnSaltoLinDoc
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
    ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal cMatCond As String, ByVal pnCar1I As Integer, ByVal pnCar1F As Integer, _
    ByVal pnCar2I As Integer, ByVal pnCar2F As Integer, ByVal pnCar3I As Integer, ByVal pnCar3F As Integer, ByVal pnCar4I As Integer, Optional ByVal psNomCmac As String = "", _
    Optional ByVal pnTipCam As Double, Optional ByVal psTitProductos As String) As String 'DAOR 20070717, se agregí pnTipCam
    'ByVal pMatCond As Variant
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaResumenSaldosCarteraPorAnalista(pMatAgencias, pdFecIni, pnMoneda, pnCar1I, pnCar1F, pnCar2I, pnCar2F, pnCar3I, pnCar3F, pnCar4I, cMatCond, pMatProd, pnTipCam)
    Set oDCred = Nothing
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "RESUMEN DE SALDOS DE CARTERA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & "Agencia     : " & R!cAgencia & lnSaltoLinDoc
        'Comentado por DAOR 20070717
        'sCadImp = sCadImp & "Moneda    : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & "Tipo Cambio : " & Str(pnTipCam) & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos   : " & psTitProductos
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
        'Comentado por DAOR 20070427
        'sCadImp = sCadImp & ImpreFormat("ANALISTA", 9) & ImpreFormat("CARTERA TOTAL", 15) & "|" & ImpreFormat("CARTERA VENCIDA 1-7 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA 8-15 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA 16-30 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA > 30 DIAS", 24) & "|" & ImpreFormat("TOTAL CART VENCIDA    ", 20) & lnSaltoLinDoc
        'sCadImp = sCadImp & Space(11) & ImpreFormat("   Nro    SALDO  ", 15) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MOR", 20) & lnSaltoLinDoc
        
        '**DAOR 20070427, Titulos de Columnas del Reporte
        sCadImp = sCadImp & ImpreFormat("ANAL", 6, 0)
        sCadImp = sCadImp & ImpreFormat("CARTERA TOTAL", 15) & "|"
        sCadImp = sCadImp & Space(16) & ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(16) & ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(16) & ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(16) & ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(16) & ImpreFormat("%MORA", 5) & "|"
        sCadImp = sCadImp & Space(16) & ImpreFormat("TOTAL", 5) & lnSaltoLinDoc
        sCadImp = sCadImp & Space(6) & ImpreFormat("   Nro    SALDO  ", 15) & "|"
        sCadImp = sCadImp & ImpreFormat("Nro   SALDO     1 - 7", 21) & "|"
        sCadImp = sCadImp & ImpreFormat("Nro   SALDO     8 -15", 21) & "|"
        sCadImp = sCadImp & ImpreFormat("Nro   SALDO     16-30", 21) & "|"
        sCadImp = sCadImp & ImpreFormat("Nro   SALDO     > 30 ", 21) & "|"
        sCadImp = sCadImp & ImpreFormat("Nro   SALDO     JUD  ", 21) & "|"
        sCadImp = sCadImp & ImpreFormat("Nro   SALDO     %MORA", 21) & lnSaltoLinDoc
        
        sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
            'FINAL DE AGENCIA
            If R!cUser = "TOTAL" Then
                sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("SUB T:", 6, 0) '
                'Comentado por DAOR 20070717
                'sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                'sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(CDbl(Format(((R!nCar1Saldo / R!nCar5Saldo) * 100), "#0.00")), 3)
                'sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 7, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(CDbl(Format(((R!nCar2Saldo / R!nCar5Saldo) * 100), "#0.00")), 3)
                'sCadImp = sCadImp & ImpreFormat(R!nCar3Nro, 8, 0) & ImpreFormat(R!nCar3Saldo, 10) & ImpreFormat(CDbl(Format(((R!nCar3Saldo / R!nCar5Saldo) * 100), "#0.00")), 4)
                'sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 7, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(CDbl(Format(((R!nCar4Saldo / R!nCar5Saldo) * 100), "#0.00")), 4)
                'sCadImp = sCadImp & ImpreFormat(CDbl(Format(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro, "#0.0000")), 9, 0) & ImpreFormat(CDbl(Format(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo, "#0.00")), 9) & ImpreFormat(CDbl(Format((((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo) / R!nCar5Saldo) * 100), "#0.00")), 3)
                'sCadImp = sCadImp & lnSaltoLinDoc
                'DAOR 20070717******************************************
                sCadImp = sCadImp & ImpreFormat(R!nCar6Nro, 5, 0) & ImpreFormat(R!nCar6Saldo, 9, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 9, 2) & ImpreFormat((R!nCar1Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 9, 2) & ImpreFormat((R!nCar2Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCar3Nro, 5, 0) & ImpreFormat(R!nCar3Saldo, 9, 2) & ImpreFormat((R!nCar3Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 9, 2) & ImpreFormat((R!nCar4Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 5, 0) & ImpreFormat(R!nCar5Saldo, 9, 2) & ImpreFormat((R!nCar5Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro + R!nCar5Nro, 6, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo, 9, 2) & ImpreFormat(((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo) / R!nCar6Saldo) * 100, 4, 2)
                '*******************************************************
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!cUser = "TOTAL" Then
                    sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("TOTAL:", 6, 0)
                    '**Comentado por DAOR 20070717
                    'sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 4, 0) & ImpreFormat(R!nCar5Saldo, 10)
                    'sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(Format((R!nCar1Saldo / R!nCar5Saldo) * 100, "#0.00"), 4)
                    'sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 7, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(Format((R!nCar2Saldo / R!nCar5Saldo) * 100, "#0.00"), 4)
                    'sCadImp = sCadImp & ImpreFormat(R!nCar3Nro, 8, 0) & ImpreFormat(R!nCar3Saldo, 10) & ImpreFormat(Format((R!nCar3Saldo / R!nCar5Saldo) * 100, "#0.00"), 4)
                    'sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(Format((R!nCar4Saldo / R!nCar5Saldo) * 100, "#0.00"), 4)
                    'sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro, 10, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo, 9) & ImpreFormat(Format(((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo) / R!nCar5Saldo) * 100, "#0.00"), 4)
                    '**DAOR 20070717******************************************
                    sCadImp = sCadImp & ImpreFormat(R!nCar6Nro, 5, 0) & ImpreFormat(R!nCar6Saldo, 9, 2)
                    sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 9, 2) & ImpreFormat((R!nCar1Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 9, 2) & ImpreFormat((R!nCar2Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & ImpreFormat(R!nCar3Nro, 5, 0) & ImpreFormat(R!nCar3Saldo, 9, 2) & ImpreFormat((R!nCar3Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 9, 2) & ImpreFormat((R!nCar4Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 5, 0) & ImpreFormat(R!nCar5Saldo, 9, 2) & ImpreFormat((R!nCar5Saldo / R!nCar6Saldo) * 100, 4, 2)
                    sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro + R!nCar5Nro, 6, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo, 9, 2) & ImpreFormat(((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo) / R!nCar6Saldo) * 100, 4, 2)
                    '*********************************************************
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & Chr$(12)
                    'Comentado por DAOR 20070717
                    'sCadImp = sCadImp & " Agencia : " & R!cAgencia & lnSaltoLinDoc
                    'sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
                    'sCadImp = sCadImp & "Productos : "
                    'sCadImp = sCadImp & lnSaltoLinDoc
                    'sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
                    'sCadImp = sCadImp & ImpreFormat("ANALISTA", 9) & ImpreFormat("CARTERA TOTAL", 15) & "|" & ImpreFormat("CARTERA VENCIDA 1-7 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA 8-15 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA 16-30 DIAS", 24) & "|" & ImpreFormat("CARTERA VENCIDA > 30 DIAS", 24) & "|" & ImpreFormat("TOTAL CART VENCIDA    ", 20) & lnSaltoLinDoc
                    'sCadImp = sCadImp & Space(11) & ImpreFormat("   Nro    SALDO  ", 15) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro      SALDO   %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MORA", 24) & "|" & ImpreFormat("Nro       SALDO  %MOR", 20) & lnSaltoLinDoc
                    'sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
                    
                    '**DAOR 20070427, Titulos de Columnas del Reporte
                    sCadImp = sCadImp & "Agencia     : " & R!cAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & "Tipo Cambio : " & Str(pnTipCam) & lnSaltoLinDoc
                    sCadImp = sCadImp & "Productos   : " & psTitProductos
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("ANAL", 6, 0)
                    sCadImp = sCadImp & ImpreFormat("CARTERA TOTAL", 15) & "|"
                    sCadImp = sCadImp & Space(16) & ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & ImpreFormat("%MORA", 5) & "|"
                    sCadImp = sCadImp & Space(16) & ImpreFormat("TOTAL", 5) & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(6) & ImpreFormat("   Nro    SALDO  ", 15) & "|"
                    sCadImp = sCadImp & ImpreFormat("Nro   SALDO     1 - 7", 21) & "|"
                    sCadImp = sCadImp & ImpreFormat("Nro   SALDO     8 -15", 21) & "|"
                    sCadImp = sCadImp & ImpreFormat("Nro   SALDO     16-30", 21) & "|"
                    sCadImp = sCadImp & ImpreFormat("Nro   SALDO     > 30 ", 21) & "|"
                    sCadImp = sCadImp & ImpreFormat("Nro   SALDO     JUD  ", 21) & "|"
                    sCadImp = sCadImp & ImpreFormat("Nro   SALDO     %MORA", 21) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(167, "-") & lnSaltoLinDoc
                End If
            End If
                'Comentado por DAOR 20070512
                'sCadImp = sCadImp & ImpreFormat(R!cUser, 9)
                'sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 5, 0) & ImpreFormat(R!nCar5Saldo, 10)
                'sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 10) & ImpreFormat(Format(((R!nCar1Saldo / R!nCar5Saldo) * 100), "#0.00"), 6)
                'sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 10) & ImpreFormat(Format(((R!nCar2Saldo / R!nCar5Saldo) * 100), "#0.00"), 7)
                'sCadImp = sCadImp & ImpreFormat(R!nCar3Nro, 5, 0) & ImpreFormat(R!nCar3Saldo, 10) & ImpreFormat(Format(((R!nCar3Saldo / R!nCar5Saldo) * 100), "#0.00"), 7)
                'sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 10) & ImpreFormat(Format(((R!nCar4Saldo / R!nCar5Saldo) * 100), "#0.00"), 6)
                'sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro, 8, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo, 9) & ImpreFormat(Format((((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo) / R!nCar5Saldo) * 100), "#0.00"), 4)
                'sCadImp = sCadImp & lnSaltoLinDoc
        
                '**DAOR 20070512
                sCadImp = sCadImp & ImpreFormat(R!cUser, 6, 0)
                sCadImp = sCadImp & ImpreFormat(R!nCar6Nro, 5, 0) & ImpreFormat(R!nCar6Saldo, 9, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro, 5, 0) & ImpreFormat(R!nCar1Saldo, 9, 2) & ImpreFormat((R!nCar1Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCar2Nro, 5, 0) & ImpreFormat(R!nCar2Saldo, 9, 2) & ImpreFormat((R!nCar2Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCar3Nro, 5, 0) & ImpreFormat(R!nCar3Saldo, 9, 2) & ImpreFormat((R!nCar3Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCar4Nro, 5, 0) & ImpreFormat(R!nCar4Saldo, 9, 2) & ImpreFormat((R!nCar4Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCar5Nro, 5, 0) & ImpreFormat(R!nCar5Saldo, 9, 2) & ImpreFormat((R!nCar5Saldo / R!nCar6Saldo) * 100, 4, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCar1Nro + R!nCar2Nro + R!nCar3Nro + R!nCar4Nro + R!nCar5Nro, 6, 0) & ImpreFormat(R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo, 9, 2) & ImpreFormat(((R!nCar1Saldo + R!nCar2Saldo + R!nCar3Saldo + R!nCar4Saldo + R!nCar5Saldo) / R!nCar6Saldo) * 100, 4, 2)
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
Dim nNumLineas As Integer

Dim scCtaCodAnt As String
Dim snDatos As String
Dim scTipoAnt As String

Dim oDCreDocT As DCredDoc
Dim rs As ADODB.Recordset
Dim bTelefono As Boolean
Dim sDireccionTrabajo As String
Dim sConvenio As String
Dim nTotal As Integer
scCtaCodAnt = ""
    
    
    If pn_ConCuotaLibre = True Then

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaMoraxAnalista_AtrasoPagoCuotaLibre(pMatAgencias, pdFecIni, pnMoneda, pMatCond, pMatProd, pMatAnalistas, pn_ConCuotaLibre)
    Set oDCred = Nothing
    
    sCadImp = ""
    nNumLineas = 0
    
    If R.RecordCount > 0 Then
        nTotal = R.RecordCount
        RaiseEvent IniciaBarra(nTotal)
        
        If pn_ConCuotaLibre = True Then
            Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "ATRASO EN PAGO DE CALENDARIO DE CUOTA LIBRE AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        ElseIf pn_ConCuotaLibre = False Then
            Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "MORA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        End If

        
            sCadImp = sCadImp & Chr$(10)
            sCadImp = sCadImp & Chr$(10)
            sCadImp = sCadImp & " Agencia : " & R!cAgencia & Chr$(10)
            sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & Chr$(10)
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
            sCadImp = sCadImp & Chr$(10)
            sCadImp = sCadImp & "ANALISTA : " & Trim(R!cAnalista) & Chr$(10)
            'sCadImp = sCadImp & String(145, "-") & Chr$(10)
            sCadImp = sCadImp & String(160, "-") & Chr$(10)
            'sCadImp = sCadImp & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("  MONTO", 13) & ImpreFormat("SALDO", 7) & ImpreFormat("CUOTA", 6, 2) & ImpreFormat("CUOTA EN", 10) & ImpreFormat("FECHA", 12, 0) & ImpreFormat("SALDO", 6, 1) & ImpreFormat("DIAS ", 6, 4) & ImpreFormat("SALDO", 6, 2) & ImpreFormat("SALDO", 12, 2) & Chr$(10)
            'sCadImp = sCadImp & ImpreFormat("                                    ", 54) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 8) & ImpreFormat("APROBADA", 9) & ImpreFormat("MORA ", 7) & ImpreFormat("VENC.", 11) & ImpreFormat("CUOTA ", 7, 0) & ImpreFormat("ATR.", 6, 4) & ImpreFormat("MORA ", 6, 2) & ImpreFormat("GASTO      ", 12, 2) & Chr$(10)
            sCadImp = sCadImp & ImpreFormat("CREDITO/CLIENTE", 25) & ImpreFormat("DIRECCION", 33) & ImpreFormat("DESEM", 5) & ImpreFormat("SALDO.K", 8) & ImpreFormat("CUOTA.APR", 10, 2) & ImpreFormat("CUOTA.MOR", 11) & ImpreFormat("FEC.VENC", 10, 0) & ImpreFormat("S.CUOTA", 7, 1) & ImpreFormat("ATRASO ", 6, 4) & ImpreFormat("S.MOR", 9, 2) & ImpreFormat("S.GASTO", 12, 2) & Chr$(10)
            'sCadImp = sCadImp & String(145, "-") & Chr$(10)
            sCadImp = sCadImp & String(160, "-") & Chr$(10)
            sCadCab = sCadImp
        End If
    
        nNumLineas = 13
        
        Do While Not R.EOF
            'FINAL DE ANALISTA
            If R!cCtaCod = "TOTAL" Then
                'sCadImp = sCadImp & String(145, "-") & Chr$(10)
                sCadImp = sCadImp & String(160, "-") & Chr$(10)
                
                'sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL ANALISTA ", 104) & ImpreFormat(Format$(R!nSaldoCuota, "#0.00"), 14, 2) & _
                        'ImpreFormat(Trim(Str(R!nDiasAtraso)), 4, 3) & ImpreFormat(Format$(R!nMora, "#0.00"), 6, 2) & ImpreFormat(Format$(R!NGasto, "#0.00"), 8, 2) & Chr$(10)
    
                sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL ANALISTA ", 138) & ImpreFormat(Format$(R!nSaldoCuota, "#0.00"), 12, 2) & _
                        ImpreFormat(Trim(Str(R!nDiasAtraso)), 4, 0) & ImpreFormat(Format$(R!nMora, "#0.00"), 6, 2) & ImpreFormat(Format$(R!NGasto, "#0.00"), 8, 2) & Chr$(10)
                        
          '      sCadImp = sCadImp & String(145, "-") & Chr$(10)
                sCadImp = sCadImp & String(160, "-") & Chr$(10)
                nNumLineas = nNumLineas + 3
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE AGENCIA
                If R!cAnalista = "TOTAL" Then
                    'sCadImp = sCadImp & String(145, "-") & Chr$(10)
                    sCadImp = sCadImp & String(160, "-") & Chr$(10)
                    'sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL AGENCIA ", 104) & ImpreFormat(Format(R!nSaldoCuota, "#0.00"), 12, 2) & _

                     'sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL AGENCIA ", 138) & ImpreFormat(Format(R!nSaldoCuota, "#0.00"), 12, 2) & _
                     'ImpreFormat(Trim(Str(R!nDiasAtraso)), 4, 3) & ImpreFormat(Format(R!nMora, "#0.00"), 6) & ImpreFormat(Format(R!NGasto, "#0.00"), 6) & Chr$(10)
                        
                     sCadImp = sCadImp & ImpreFormat("*** SUB TOTAL AGENCIA ", 138) & ImpreFormat(Format$(R!nSaldoCuota, "#0.00"), 12, 2) & _
                     ImpreFormat(Trim(Str(R!nDiasAtraso)), 4, 0) & ImpreFormat(Format$(R!nMora, "#0.00"), 6, 2) & ImpreFormat(Format$(R!NGasto, "#0.00"), 8, 2) & Chr$(10)
                       
                    'sCadImp = sCadImp & String(145, "-") & Chr$(10)
                    sCadImp = sCadImp & String(160, "-") & Chr$(10)
                    nNumLineas = nNumLineas + 3
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL FINAL
                    If R!cAgencia = "TOTAL" Then
                        ' sCadImp = sCadImp & String(145, "-") & Chr$(10)
                        sCadImp = sCadImp & String(160, "-") & Chr$(10)
                        
                        'sCadImp = sCadImp & ImpreFormat("*** TOTAL ", 104) & ImpreFormat(Format(R!nSaldoCuota, "#0.00"), 12) & _

    '                    sCadImp = sCadImp & ImpreFormat("*** TOTAL ", 138) & ImpreFormat(Format(R!nSaldoCuota, "#0.00"), 12) & _
     '                   ImpreFormat(Trim(Str(R!nDiasAtraso)), 4, 3) & ImpreFormat(Format(R!nMora, "#0.00"), 6) & ImpreFormat(Format(R!NGasto, "#0.00"), 6) & Chr$(10)
                        
                        sCadImp = sCadImp & ImpreFormat("*** TOTAL ", 138) & ImpreFormat(Format$(R!nSaldoCuota, "#0.00"), 12, 2) & _
                        ImpreFormat(Trim(Str(R!nDiasAtraso)), 4, 0) & ImpreFormat(Format$(R!nMora, "#0.00"), 6, 2) & ImpreFormat(Format$(R!NGasto, "#0.00"), 8, 2) & Chr$(10)
                        
                        'sCadImp = sCadImp & String(145, "-") & Chr$(10)
                        sCadImp = sCadImp & String(160, "-") & Chr$(10)
                        nNumLineas = nNumLineas + 3
                        Exit Do
                    Else
                        sCadCab = sCadCab & Chr$(10)
                        sCadImp = sCadImp & " Agencia : " & R!cAgencia & Chr$(10)
                        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & Chr$(10)
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
                        sCadCab = sCadCab & Chr$(10)
                        sCadCab = sCadCab & "ANALISTA : " & Trim(R!cAnalista) & Chr$(10)
                        sCadCab = sCadCab & Chr$(10)
                        'sCadCab = sCadCab & String(145, "-") & Chr$(10)
                        sCadCab = sCadCab & String(160, "-") & Chr$(10)
                        'sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DIRECCION", 34) & ImpreFormat("  MONTO", 13) & ImpreFormat("SALDO", 7) & ImpreFormat("CUOTA", 6, 2) & ImpreFormat("CUOTA EN", 10) & ImpreFormat("FECHA", 12, 0) & ImpreFormat("SALDO", 6, 1) & ImpreFormat("DIAS ", 6, 4) & ImpreFormat("SALDO", 6, 2) & ImpreFormat("SALDO", 12, 2) & Chr$(10)
                        'sCadCab = sCadCab & ImpreFormat("                                    ", 54) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 8) & ImpreFormat("APROBADA", 9) & ImpreFormat("MORA ", 7) & ImpreFormat("VENC.", 11) & ImpreFormat("CUOTA ", 7, 0) & ImpreFormat("ATR.", 6, 4) & ImpreFormat("MORA ", 6, 2) & ImpreFormat("GASTO      ", 12, 2) & Chr$(10)
                        
                        sCadImp = sCadImp & ImpreFormat("CREDITO", 18) & ImpreFormat("CLIENTE", 30) & ImpreFormat("DIRECCION", 28) & ImpreFormat("MONTO", 11) & ImpreFormat("SALDO", 8) & ImpreFormat("CUOTA", 6, 2) & ImpreFormat("CUOTA EN", 11) & ImpreFormat("FECHA", 10, 0) & ImpreFormat("SALDO", 7, 1) & ImpreFormat("DIAS ", 6, 4) & ImpreFormat("SALDO", 6, 2) & ImpreFormat("SALDO", 12, 2) & Chr$(10)
                        sCadImp = sCadImp & ImpreFormat("                                    ", 82) & ImpreFormat("DESEM.", 8) & ImpreFormat("CAPITAL", 8) & ImpreFormat("APROB.", 5) & ImpreFormat("MORA ", 7) & ImpreFormat("VENC.", 11) & ImpreFormat("CUOTA ", 7, 0) & ImpreFormat("ATR.", 6, 4) & ImpreFormat("MORA ", 6, 2) & ImpreFormat("GASTO      ", 12, 2) & Chr$(10)
                        
                        'sCadCab = sCadCab & String(145, "-") & Chr$(10)
                        sCadCab = sCadCab & String(160, "-") & Chr$(10)
                        sCadImp = sCadImp & sCadCab
                        nNumLineas = nNumLineas + 10
                    End If
                Else
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
                        sCadCab = sCadCab & Chr$(10)
                        sCadCab = sCadCab & "ANALISTA : " & Trim(R!cAnalista) & Chr$(10)
                        sCadCab = sCadCab & Chr$(10)
                        'sCadCab = sCadCab & String(145, "-") & Chr$(10)
                        sCadCab = sCadCab & String(160, "-") & Chr$(10)
                        
                        'sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DIRECCION", 34) & ImpreFormat("  MONTO", 13) & ImpreFormat("SALDO", 7) & ImpreFormat("CUOTA", 6, 2) & ImpreFormat("CUOTA EN", 10) & ImpreFormat("FECHA", 12, 0) & ImpreFormat("SALDO", 6, 1) & ImpreFormat("DIAS ", 6, 4) & ImpreFormat("SALDO", 6, 2) & ImpreFormat("SALDO", 12, 2) & Chr$(10)
                        'sCadCab = sCadCab & ImpreFormat("                                    ", 54) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 8) & ImpreFormat("APROBADA", 9) & ImpreFormat("MORA ", 7) & ImpreFormat("VENC.", 11) & ImpreFormat("CUOTA ", 7, 0) & ImpreFormat("ATR.", 6, 4) & ImpreFormat("MORA ", 6, 2) & ImpreFormat("GASTO      ", 12, 2) & Chr$(10)
                        
                        sCadImp = sCadImp & ImpreFormat("CREDITO", 18) & ImpreFormat("CLIENTE", 32) & ImpreFormat("DIRECCION", 31) & ImpreFormat("MONTO", 11) & ImpreFormat("SALDO", 8) & ImpreFormat("CUOTA", 6, 2) & ImpreFormat("CUOTA EN", 11) & ImpreFormat("FECHA", 10, 0) & ImpreFormat("SALDO", 7, 1) & ImpreFormat("DIAS ", 6, 4) & ImpreFormat("SALDO", 6, 2) & ImpreFormat("SALDO", 12, 2) & Chr$(10)
                        sCadImp = sCadImp & ImpreFormat("                                    ", 82) & ImpreFormat("DESEMB.", 13) & ImpreFormat("CAPITAL", 8) & ImpreFormat("APROB.", 9) & ImpreFormat("MORA ", 7) & ImpreFormat("VENC.", 11) & ImpreFormat("CUOTA ", 7, 0) & ImpreFormat("ATR.", 6, 4) & ImpreFormat("MORA ", 6, 2) & ImpreFormat("GASTO      ", 12, 2) & Chr$(10)
                        
                        'sCadCab = sCadCab & String(145, "-") & Chr$(10)
                        sCadCab = sCadCab & String(160, "-") & Chr$(10)
                        sCadImp = sCadImp & sCadCab
                        nNumLineas = nNumLineas + 7
                End If
            End If
    
            If Not IsNull(R!ndatos) Then
                If R!cTipo = "A" Then
                    If R!cCtaCod <> scCtaCodAnt Then
                        sCadImp = sCadImp & Chr(10)
                        sCadImp = sCadImp & String(160, "-") & Chr(10)
                        'insertando los datos del titular del credito
                        'Modificado por LMMD
                        bTelefono = False
                        Set oDCreDocT = New DCredDoc
                        sDireccionTrabajo = oDCreDocT.ObtenerDireccTrabajoTitular(R!cCtaCod)
                        sConvenio = oDCreDocT.ObtenerConvenio(R!cCtaCod)
                        Set rs = oDCreDocT.ObtenerTelefonosTitular(R!cCtaCod)
                        Set oDCreDocT = Nothing
                        sDireccionTrabajo = "Dir.Trabaj: " & sDireccionTrabajo
                        If Not rs.EOF And Not rs.BOF Then
                            If Not IsNull(rs!cPersTelefono) Then
                                bTelefono = True
                                If Len(Trim(rs!cPersTelefono)) > 0 Then
                                    sTelefono = "Telefono: " & Trim(rs!cPersTelefono)
                                Else
                                    sTelefono = "Telefono: "
                                End If
                            ElseIf Not IsNull(Trim(rs!cPersTelefono2)) Then
                                bTelefono = True
                                If Len(rs!cPersTelefono2) > 0 Then
                                    sTelefono = "Telefono: " & Trim(rs!cPersTelefono2)
                                End If
                            Else
                                sTelefono = "Telefono:"
                            End If
                        Else
                            sTelefono = "Telefono:"
                        End If
                        Set rs = Nothing
                        
                        
    '                    sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 18) & ImpreFormat(R!nDatos, 118, 2) & _
    '                          ImpreFormat(Format$(R!nSaldoCuota, "#0.00"), 10, 2) & ImpreFormat(R!nDiasAtraso, 4, 0) & ImpreFormat(Format$(R!nMora, "#0.00"), 9, 2) & ImpreFormat(Format$(R!NGasto, "#0.00"), 6, 2) & Chr$(10)
                        sCadImp = sCadImp & ImpreFormat(R!ndatos, 110, 2) & _
                                  ImpreFormat(Format$(R!nSaldoCuota, "#0.00"), 10, 2) & ImpreFormat(R!nDiasAtraso, 6, 0) & ImpreFormat(Format$(R!nMora, "#0.00"), 9, 2) & ImpreFormat(Format$(R!NGasto, "#0.00"), 6, 2) & Chr$(10)
                        sCadImp = sCadImp & Space(37) & sDireccionTrabajo & "  " & sTelefono & "  " & sConvenio & Chr(10)
                        sCadImp = sCadImp & " Nro Credito:" & ImpreFormat(R!cCtaCod, 18) & Chr(10)
                    Else
                        snDatos = Mid(R!ndatos, 67, Len(R!ndatos))
                        sCadImp = sCadImp & ImpreFormat(" ", 78) & ImpreFormat(snDatos, 51, 2) & _
                              ImpreFormat(Format$(R!nSaldoCuota, "#0.00"), 6, 2) & ImpreFormat(R!nDiasAtraso, 3, 0) & ImpreFormat(Format$(R!nMora, "#0.00"), 6, 2) & ImpreFormat(Format$(R!NGasto, "#0.00"), 5, 2) & Chr$(10)
                    End If
                                  
                Else
                    If R!cTipo <> scTipoAnt Then
                       sCadImp = sCadImp & ImpreFormat("     AVAL", 18) & ImpreFormat(R!ndatos, 60, 2) & Chr$(10)
                    Else
                       sCadImp = sCadImp & ImpreFormat("         ", 18) & ImpreFormat(R!ndatos, 60, 2) & Chr$(10)
                    End If
                End If
                scCtaCodAnt = R!cCtaCod
                scTipoAnt = R!cTipo
                nNumLineas = nNumLineas + 1
            End If
            
            If nNumLineas > 54 Then
                sCadImp = sCadImp & Chr$(12)
                
                If pn_ConCuotaLibre = True Then
                    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "ATRASO EN PAGO DE CALENDARIO DE CUOTA LIBRE AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
                ElseIf pn_ConCuotaLibre = False Then
                    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "MORA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
                End If
                
                sCadCab = Chr$(10)
                sCadCab = sCadCab & Chr$(10)
                sCadCab = sCadCab & " Agencia : " & R!cAgencia & Chr$(10)
                sCadCab = sCadCab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & Chr$(10)
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
                sCadCab = sCadCab & Chr$(10)
                sCadCab = sCadCab & "ANALISTA : " & Trim(R!cAnalista) & Chr$(10)
                sCadCab = sCadCab & Chr$(10)
    '            sCadCab = sCadCab & String(145, "-") & Chr$(10)
                sCadCab = sCadCab & String(160, "-") & Chr$(10)
                'sCadCab = sCadCab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("DIRECCION", 34) & ImpreFormat("  MONTO", 13) & ImpreFormat("SALDO", 7) & ImpreFormat("CUOTA", 6, 2) & ImpreFormat("CUOTA EN", 10) & ImpreFormat("FECHA", 12, 0) & ImpreFormat("SALDO", 6, 1) & ImpreFormat("DIAS ", 6, 4) & ImpreFormat("SALDO", 6, 2) & ImpreFormat("SALDO", 12, 2) & Chr$(10)
                'sCadCab = sCadCab & ImpreFormat("                                    ", 54) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 8) & ImpreFormat("APROBADA", 9) & ImpreFormat("MORA ", 7) & ImpreFormat("VENC.", 11) & ImpreFormat("CUOTA ", 7, 0) & ImpreFormat("ATR.", 6, 4) & ImpreFormat("MORA ", 6, 2) & ImpreFormat("GASTO      ", 12, 2) & Chr$(10)
               
                sCadImp = sCadImp & ImpreFormat("CREDITO/CLIENTE", 25) & ImpreFormat("DIRECCION", 33) & ImpreFormat("DESEM", 10) & ImpreFormat("SALDO.K", 8) & ImpreFormat("CUOTA.APR", 10, 2) & ImpreFormat("CUOTA.MOR", 11) & ImpreFormat("FEC.VENC", 10, 0) & ImpreFormat("S.CUOTA", 7, 1) & ImpreFormat("ATRASO ", 6, 4) & ImpreFormat("S.MOR", 9, 2) & ImpreFormat("S.GASTO", 12, 2) & Chr$(10)
                
                'sCadCab = sCadCab & String(145, "-") & Chr$(10)
                sCadCab = sCadCab & String(160, "-") & Chr$(10)
                sCadImp = sCadImp & sCadCab
                nNumLineas = 14
            End If
            RaiseEvent ProgresoBarra(I, "MORA POR ANALISTA", "PROCESANDO...")
            I = I + 1
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        RaiseEvent FinalizaBarra
        ImprimeMoraXAnalista_AtrasoPagoCuotaLibre = sCadImp
  Else
   'ImprimeMoraXAnalista_AtrasoPagoCuotaLibre = ImprimeMoraXAnalitaNew(pMatAgencias, pdFecIni, CStr(pnMoneda), psCodUser, pdfecSis, psNomAge, pMatCond, pMatprod, pMatAnalistas, psNomCmac)
  End If
End Function

Public Function ImpreMoraxAnalista(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, _
ByVal pnMoneda As Integer, ByVal psCodUser As String, ByVal pdFecSis As Date, _
ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, _
ByVal pMatAnalistas As Variant, Optional ByVal pn_CuotaLibre As Boolean = False, _
Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim nNumLineas As Integer

Dim scCtaCodAnt As String
Dim snDatos As String
Dim scTipoAnt As String

Dim oDCreDocT As DCredDoc
Dim rs As ADODB.Recordset
Dim bTelefono As Boolean
Dim sDireccionTrabajo As String
Dim sConvenio As String
Dim nTotal As Integer
scCtaCodAnt = ""


    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaMoraxAnalista_AtrasoPagoCuotaLibre(pMatAgencias, pdFecIni, pnMoneda, pMatCond, pMatProd, pMatAnalistas, pn_CuotaLibre)
    Set oDCred = Nothing
    
    sCadImp = ""
    nNumLineas = 0


    If R.RecordCount > 0 Then
        nTotal = R.RecordCount
        RaiseEvent IniciaBarra(nTotal)
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "MORA POR ANALISTA AL " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True)
        
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & " Agencia : " & R!cAgencia & Chr$(10)
        sCadImp = sCadImp & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & Chr$(10)
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
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & "ANALISTA : " & Trim(R!cAnalista) & Chr$(10)
        sCadImp = sCadImp & String(175, "-") & Chr$(10)
        
        sCadImp = sCadImp & ImpreFormat("Codigo", 20) & ImpreFormat("Cliente", 15) & ImpreFormat("Prestamo", 10) & ImpreFormat("Sal.Cap.Tot", 12) & ImpreFormat("Cuot#", 5) & ImpreFormat("Cuot.Mora", 6) & ImpreFormat("S.Cuota", 6) & ImpreFormat("Atraso", 7) & ImpreFormat("Cobrar", 8) & Chr(10)
        sCadImp = sCadImp & String("165", "-") & Chr(10)
        
        Do Until R.EOF
            If Not IsNull(R!ndatos) Then
                
            End If
            R.MoveNext
        Loop
    End If
    
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
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!cpersnombre, 25) & ImpreFormat(R!nSaldo, 10) & _
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
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS ANULADOS DEL " & Format(pdFecIni, "dd/mm/yyyy") & "  AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
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
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!cpersnombre, 34) & ImpreFormat(R!nMonto, 10) & ImpreFormat(R!cConsDescripcion, 20) & ImpreFormat(Format(R!dPrdEstado, "dd/mm/yyyy"), 20) & ImpreFormat(R!cUser, 10)
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
        sCadFormato = Replace(sCadFormato, "<<NOMBRE>>", PstaNombre(R!cpersnombre), , 1, vbTextCompare)
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
        sCadFormato = Replace(sCadFormato, "<<NOMBRE>>", PstaNombre(R!cpersnombre), , 1, vbTextCompare)
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
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & Space(1) & ImpreFormat(PstaNombre(R!cpersnombre), 34) & Space(1) & ImpreFormat(R!cPersDireccDomicilio, 20) & Space(1) & ImpreFormat(R!Zona, 20) & Space(1) & ImpreFormat(R!nMonto, 10) & Space(1) & ImpreFormat(R!nSaldo, 10) & Space(1) & ImpreFormat(IIf(IsNull(R!nColocNota), 0, R!nColocNota), 5, 0) & Space(1) & ImpreFormat(R!cUser, 5, 0) & Space(1) & ImpreFormat(R!nNroProxCuota & "/" & R!nCuotas, 5, 0) & lnSaltoLinDoc
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
Dim nLineas As Integer
Dim sCadTempo As String
Dim nLineasTempo As Integer
Dim oNCred As NCredito
Dim nTempoDeudaFecha As Double
Dim MatCalend As Variant
Dim MatCalendPar As Variant

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
        
        nLineas = 4
        
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & Chr$(10)
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & Chr$(10)
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & Chr$(10)
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(230, "-") & Chr$(10)
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & _
            ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
            ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
            ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
            ImpreFormat("INTERES", 8) & ImpreFormat("DEUDA TOT", 12, 0) & ImpreFormat("CAP VENC", 7, 0) & _
            ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
            ImpreFormat("REFINANC", 12) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(230, "-") & Chr$(10)
    End If
    
    nLineas = nLineas + 11
    sCadTempo = ""
    nLineasTempo = 0
    Do While Not R.EOF
    'FINAL DE Financiamiento
        If Trim(R!cCtaCod) = "TOTAL" Then
            
            sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
            sCadTempo = sCadTempo & lsTab & ImpreFormat("*** SUB TOTAL FINANCIAMIENTO ", 42) & _
                        ImpreFormat("CASOS : " & R!nCasos, 45) & ImpreFormat(R!nMontoApr, 10, , True) & _
                        ImpreFormat(R!nSaldoCap, 10, , True) & ImpreFormat(Format(R!IntDeveng, "#0.00"), 47, , True) & _
                        ImpreFormat(Format(R!nDeudaTotal, "#0.00"), 10, , True) & ImpreFormat(Format(R!nCapVenc, "#0.00"), 8, , True) & _
                        ImpreFormat(Format(R!nCapXVencer, "#0.00"), 8, , True) & ImpreFormat(R!nMontoGarHip, 8, , True) & _
                        ImpreFormat(R!nMontoOtrasGarant, 10, , True) & Chr$(10)
            sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
            sCadTempo = sCadTempo & Chr$(10)
            
            nLineas = nLineas + 4
            nLineasTempo = nLineasTempo + 4
            If nLineas < 55 Then
                sCadImp = sCadImp & sCadTempo
                 sCadTempo = ""
                 nLineasTempo = 0
            End If
            
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE PLAZO
            If R!sFuenteFinan = "TOTAL" Then
                sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                sCadTempo = sCadTempo & lsTab & ImpreFormat("*** SUB TOTAL PLAZO ", 42) & _
                ImpreFormat("CASOS : " & R!nCasos, 45) & ImpreFormat(R!nMontoApr, 10, , True) & _
                ImpreFormat(R!nSaldoCap, 10, , True) & ImpreFormat(Format(R!IntDeveng, "#0.00"), 47, , True) & _
                ImpreFormat(Format(R!nDeudaTotal, "#0.00"), 10, , True) & ImpreFormat(Format(R!nCapVenc, "#0.00"), 8, , True) & _
                ImpreFormat(Format(R!nCapXVencer, "#0.00"), 8, , True) & ImpreFormat(R!nMontoGarHip, 8, , True) & _
                ImpreFormat(R!nMontoOtrasGarant, 10, , True) & Chr$(10)
                sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                sCadTempo = sCadTempo & Chr$(10)
                nLineas = nLineas + 4
                nLineasTempo = nLineasTempo + 4
                If nLineas < 55 Then
                    sCadImp = sCadImp & sCadTempo
                     sCadTempo = ""
                     nLineasTempo = 0
                End If
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If R!sPlazo = "TOTAL" Then
                    sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                    sCadTempo = sCadTempo & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 42) & _
                    ImpreFormat("CASOS : " & R!nCasos, 45) & ImpreFormat(R!nMontoApr, 10, , True) & _
                    ImpreFormat(R!nSaldoCap, 10, , True) & ImpreFormat(Format(R!IntDeveng, "#0.00"), 47, , True) & _
                    ImpreFormat(Format(R!nDeudaTotal, "#0.00"), 10, , True) & ImpreFormat(Format(R!nCapVenc, "#0.00"), 8, , True) & _
                    ImpreFormat(Format(R!nCapXVencer, "#0.00"), 8, , True) & ImpreFormat(R!nMontoGarHip, 8, , True) & _
                    ImpreFormat(R!nMontoOtrasGarant, 10, , True) & Chr$(10)
                    sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                    sCadTempo = sCadTempo & Chr$(10)
                    nLineas = nLineas + 4
                    nLineasTempo = nLineasTempo + 4
                    If nLineas < 55 Then
                        sCadImp = sCadImp & sCadTempo
                         sCadTempo = ""
                         nLineasTempo = 0
                    End If
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL DE AGENCIAS
                    If R!sProducto = "TOTAL" Then
                        sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                        sCadTempo = sCadTempo & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 42) & _
                        ImpreFormat("CASOS : " & R!nCasos, 45) & ImpreFormat(R!nMontoApr, 10, , True) & _
                        ImpreFormat(R!nSaldoCap, 10, , True) & ImpreFormat(Format(R!IntDeveng, "#0.00"), 47, , True) & _
                        ImpreFormat(Format(R!nDeudaTotal, "#0.00"), 10, , True) & ImpreFormat(Format(R!nCapVenc, "#0.00"), 8, , True) & _
                        ImpreFormat(Format(R!nCapXVencer, "#0.00"), 8, , True) & ImpreFormat(R!nMontoGarHip, 8, , True) & _
                        ImpreFormat(R!nMontoOtrasGarant, 10, , True) & Chr$(10)
                        sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                        sCadTempo = sCadTempo & Chr$(10)
                        nLineas = nLineas + 4
                        nLineasTempo = nLineasTempo + 4
                        If nLineas < 55 Then
                            sCadImp = sCadImp & sCadTempo
                             sCadTempo = ""
                             nLineasTempo = 0
                        End If
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'TOTAL FINAL
                        If R!sAgencia = "TOTAL" Then
                            sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                            sCadTempo = sCadTempo & lsTab & ImpreFormat("*** TOTAL ", 42) & _
                            ImpreFormat("CASOS : " & R!nCasos, 45) & ImpreFormat(R!nMontoApr, 10, , True) & _
                            ImpreFormat(R!nSaldoCap, 10, , True) & ImpreFormat(Format(R!IntDeveng, "#0.00"), 47, , True) & _
                            ImpreFormat(Format(R!nDeudaTotal, "#0.00"), 10, , True) & ImpreFormat(Format(R!nCapVenc, "#0.00"), 8, , True) & _
                            ImpreFormat(Format(R!nCapXVencer, "#0.00"), 8, , True) & ImpreFormat(R!nMontoGarHip, 8, , True) & _
                            ImpreFormat(R!nMontoOtrasGarant, 10, , True) & Chr$(10)
                            sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                            sCadTempo = sCadTempo & Chr$(10)
                            nLineas = nLineas + 4
                            nLineasTempo = nLineasTempo + 4
                            If nLineas < 55 Then
                                sCadImp = sCadImp & sCadTempo
                                 sCadTempo = ""
                                 nLineasTempo = 0
                            End If
                            Exit Do
                        Else
                            nPuntPag = 0
                            
                            sCadTempo = sCadTempo & Chr$(10)
                            sCadTempo = sCadTempo & lsTab & "AGENCIA : " & Trim(R!sAgencia) & Chr$(10)
                            sCadTempo = sCadTempo & lsTab & "PRODUCTO : " & Trim(R!sProducto) & Chr$(10)
                            sCadTempo = sCadTempo & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
                            sCadTempo = sCadTempo & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                            sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                            sCadTempo = sCadTempo & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & _
                                ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                                ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                                ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                                ImpreFormat("INTERES", 8) & ImpreFormat("DEUDA TOT", 12, 0) & ImpreFormat("CAP VENC", 7, 0) & _
                                ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                                ImpreFormat("REFINANC", 12) & Chr$(10)
                            sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                            nLineas = nLineas + 8
                            nLineasTempo = nLineasTempo + 8
                            If nLineas < 55 Then
                                sCadImp = sCadImp & sCadTempo
                                sCadTempo = ""
                                nLineasTempo = 0
                            End If
                            
                        End If
                    Else
                        sCadTempo = sCadTempo & lsTab & "PRODUCTO : " & Trim(R!sProducto) & Chr$(10)
                        sCadTempo = sCadTempo & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
                        sCadTempo = sCadTempo & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                        sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                        sCadTempo = sCadTempo & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & _
                            ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                            ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                            ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                            ImpreFormat("INTERES", 8) & ImpreFormat("DEUDA TOT", 12, 0) & ImpreFormat("CAP VENC", 7, 0) & _
                            ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                            ImpreFormat("REFINANC", 12) & Chr$(10)
                        sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                        nLineas = nLineas + 6
                        nLineasTempo = nLineasTempo + 6
                        If nLineas < 55 Then
                            sCadImp = sCadImp & sCadTempo
                             sCadTempo = ""
                             nLineasTempo = 0
                        End If
                    End If
                Else
                    sCadTempo = sCadTempo & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
                    sCadTempo = sCadTempo & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                    sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                    sCadTempo = sCadTempo & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & _
                        ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                        ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                        ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                        ImpreFormat("INTERES", 8) & ImpreFormat("DEUDA TOT", 12, 0) & ImpreFormat("CAP VENC", 7, 0) & _
                        ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                        ImpreFormat("REFINANC", 12) & Chr$(10)
                    sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                    nLineas = nLineas + 5
                    nLineasTempo = nLineasTempo + 5
                     If nLineas < 55 Then
                        sCadImp = sCadImp & sCadTempo
                         sCadTempo = ""
                         nLineasTempo = 0
                    End If
                End If
            Else
                sCadTempo = sCadTempo & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                    sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                    sCadTempo = sCadTempo & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & _
                        ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                        ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                        ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                        ImpreFormat("INTERES", 8) & ImpreFormat("DEUDA TOT", 12, 0) & ImpreFormat("CAP VENC", 7, 0) & _
                        ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                        ImpreFormat("REFINANC", 12) & Chr$(10)
                    sCadTempo = sCadTempo & lsTab & String(230, "-") & Chr$(10)
                    nLineas = nLineas + 4
                    nLineasTempo = nLineasTempo + 4
                    If nLineas < 55 Then
                        sCadImp = sCadImp & sCadTempo
                        sCadTempo = ""
                        nLineasTempo = 0
                    End If
            End If
        End If
        
                Set oNCred = New NCredito
                MatCalend = oNCred.RecuperaMatrizCalendarioPendiente(R!cCtaCod)
                MatCalendPar = oNCred.RecuperaMatrizCalendarioPendiente(R!cCtaCod, True)
                MatCalend = UnirMatricesMiViviendaAmortizacion(MatCalend, MatCalendPar)
                nTempoDeudaFecha = oNCred.MatrizDeudaAlaFecha(R!cCtaCod, MatCalend, pdFecIni)
                Set oNCred = Nothing

            sCadTempo = sCadTempo & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(PstaNombre(R!sTitular), 25) & _
                                    ImpreFormat(Mid(IIf(IsNull(R!sInstitucion), ".", R!sInstitucion), 1, 22), 23) & _
                                    ImpreFormat(Format(R!dVigencia, "dd/mm/yyyy"), 12, 0) & ImpreFormat(R!sPlazoCuota, 5, 0) & ImpreFormat(R!nMontoApr, 10, , True) & ImpreFormat(R!nSaldoCap, 10, , True) _
                                    & ImpreFormat(R!nCuotaMora, 8, 0, True) & ImpreFormat(R!nDiasAtraso, 8, 0) & ImpreFormat(R!nMontoCuota, 8, , True) & ImpreFormat(Format(R!dFecUtlPago, "dd/mm/yyyy"), 12) _
                                    & ImpreFormat(CDbl(Format(R!IntDeveng, "#0.00")), 6, 2, True) & ImpreFormat(Format(nTempoDeudaFecha, "#0.00"), 10, , True) & ImpreFormat(IIf(IsNull(R!nCapVenc), 0, Format(R!nCapVenc, "#0.00")), 8, , True) _
                                    & ImpreFormat(Format(R!nCapXVencer, "#0.00"), 8, , True) & ImpreFormat(R!nMontoGarHip, 8, , True) & ImpreFormat(R!nMontoOtrasGarant, 10, , True) & ImpreFormat(R!cEstadoDesc, 12) & Chr$(10)
            nLineas = nLineas + 1
            nLineasTempo = nLineasTempo + 1
            If nLineas < 55 Then
                sCadImp = sCadImp & sCadTempo
                sCadTempo = ""
                nLineasTempo = 0
            End If
            If nLineas >= 55 Then
                sCadImp = sCadImp & Chr$(12)
                If InStr(1, sCadTempo, "AGENCIA :") = 0 Then
                    
                    sCadImp = sCadImp & Chr$(10)
                    sCadImp = sCadImp & Chr$(10)
                    sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & Chr$(10)
                    sCadImp = sCadImp & Chr$(10)
                    sCadImp = sCadImp & Chr$(10)
                    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & Chr$(10)
                    sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & Chr$(10)
                    sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                    sCadImp = sCadImp & lsTab & String(230, "-") & Chr$(10)
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 30) & _
                        ImpreFormat("INSTITUCION", 16) & ImpreFormat("F. DESEMB.", 10) & ImpreFormat("PLAZO", 5) & _
                        ImpreFormat("MONTO APROB", 13) & ImpreFormat("S. CAPITAL", 8) & ImpreFormat("C. MORA", 7) & _
                        ImpreFormat("D. MORA", 8) & ImpreFormat("M. CUOTA", 7) & ImpreFormat("ULT. FEC PAGO", 12) & _
                        ImpreFormat("INTERES", 8) & ImpreFormat("DEUDA TOT", 12, 0) & ImpreFormat("CAP VENC", 7, 0) & _
                        ImpreFormat("CAP X VENC", 8) & ImpreFormat("GAR. HIP", 9) & ImpreFormat("OTRAS GAR", 12) & _
                        ImpreFormat("REFINANC", 12) & Chr$(10)
                    sCadImp = sCadImp & lsTab & String(230, "-") & Chr$(10)
                    sCadImp = sCadImp & sCadTempo
                 Else
                    sCadImp = sCadImp & sCadTempo
                End If
                nLineas = nLineasTempo + 12
                sCadTempo = ""
                nLineasTempo = 0
            End If
            
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
    Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "", Optional ByVal MatInstitucion As Variant = "") As String

 Dim lnMoraCalc As Double
 Dim lnDias As Long
 Dim lnItfCal As Double
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
Dim lnTotal As Long
    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosXInstitucion(pMatAgencias, pnMoneda, pdFecIni, pnTipoOrden, pnIncluyeMora, MatInstitucion)
    Set oDCred = Nothing
    
    Set oGastos = New NGasto
    Set oNegCredito = New NCredito
        
    If R.RecordCount > 0 Then
        lnTotal = R.RecordCount
        
        RaiseEvent IniciaBarra(lnTotal)
        
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "LISTADO PARA DESCUENTO POR PLANILLA AL : " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True, psCodRepo)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
        
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 25) & ImpreFormat("COD. MODULAR", 15) & _
            ImpreFormat("DESEMBOLSO", 15) & ImpreFormat("SALDO CAP", 10) & ImpreFormat("C. PEND", 7) & _
            ImpreFormat("MONTO.CUO", 9) & ImpreFormat("SAL.CUOTA", 8) & ImpreFormat("ITF", 6) & ImpreFormat("TOTAL+ITF", 10) & ImpreFormat("FEC. PAGO", 10) & lnSaltoLinDoc
            
        sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE INSTITUCION
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL INSTITUCION ", 42) & ImpreFormat("CASOS : " & R!nCasos, 39) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuota, 16) & ImpreFormat(R!nSaldoCuota, 16) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE AGENCIA
            If R!cInstitucion = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 42) & ImpreFormat("CASOS : " & R!nCasos, 39) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuota, 16) & ImpreFormat(R!nSaldoCuota, 16) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'TOTAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat("CASOS : " & R!nCasos, 39) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuota, 16) & ImpreFormat(R!nSaldoCuota, 16) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
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
                    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
                    'scadImp = scadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("COD. MODULAR", 20) & _
                        ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("APROBADO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("C. PEND", 5) & _
                        ImpreFormat("CUOTA", 12) & ImpreFormat("SALDO CUOTA", 12) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
                    
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 25) & ImpreFormat("COD. MODULAR", 15) & _
                            ImpreFormat("DESEMBOLSO", 15) & ImpreFormat("SALDO CAP", 10) & ImpreFormat("C. PEND", 7) & _
                            ImpreFormat("MONTO.CUO", 8) & ImpreFormat("SAL.CUOTA", 8) & ImpreFormat("ITF", 8) & ImpreFormat("TOTAL+ITF", 12) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
                            
                                      sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
               End If
            Else
                sCadImp = sCadImp & Chr(12)
                sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
                
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 25) & ImpreFormat("COD. MODULAR", 15) & _
                        ImpreFormat("DESEMBOLSO", 15) & ImpreFormat("SALDO CAP", 10) & ImpreFormat("C. PEND", 7) & _
                        ImpreFormat("MONTO.CUO", 8) & ImpreFormat("SAL.CUOTA", 8) & ImpreFormat("ITF", 8) & ImpreFormat("TOTAL+ITF", 12) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
                        
                sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
                'scadImp = scadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("COD. MODULAR", 20) & _
                ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("APROBADO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("C. PEND", 5) & _
                        ImpreFormat("CUOTA", 12) & ImpreFormat("SALDO CUOTA", 12) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
                        
                If R!cCtaCod <> "Total" And R!cInstitucion <> "Total" Then
                   ' Dim lnMoraCalc As Double
                   ' Dim lnDias As Long
                    'Dim lnItfCal As Double
                
                    If pdFecIni > pdFecSis And R!nDiasAtraso > 0 Then
                        lnDias = DateDiff("d", pdFecSis, pdFecIni)
                        lnMoraCalc = Round(lnDias * R!nCapPend * (R!nTasMor / 100), 2)
                        lnItfCal = CalculoSinRedondeo(lnMoraCalc)
                        If R!cCtaCod = "108013011000002984" Then Stop
                        Else
                            lnMoraCalc = 0
                            lnItfCal = 0
                        End If
                
                        MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(R!cCtaCod)
                        MatGastosCancelacion = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosCancel, gdFecSis, R!cCtaCod, 1, "CA", , , R!nMontoApr, oNegCredito.MatrizMontoCapitalAPagar(MatCalend, gdFecSis), MatCalend(0, 2), , , , , R!nDiasAtraso)
                        nMontoGastoCargo = MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("CA", "PA", ""))
                
                        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(PstaNombre(R!sTitular), 25) _
                                        & ImpreFormat(R!cCodModular, 15) & ImpreFormat(Format(R!dFecDesemb, "dd/mm/yyyy"), 12) _
                                        & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!sCuotaPend, 6, 4) & ImpreFormat(R!nCuota + nMontoGastoCargo + lnMoraCalc + lnItfCal, 6) & ImpreFormat(R!nSaldoCuota + nMontoGastoCargo + lnMoraCalc + lnItfCal, 6) _
                                        & ImpreFormat(R!nITFSalCuota + lnItfCal, 6) & ImpreFormat(R!nSaldoCuota + nMontoGastoCargo + lnMoraCalc + R!nITFSalCuota + lnItfCal, 8) & ImpreFormat(Format(R!dFecPago, "dd/mm/yyyy"), 12) & lnSaltoLinDoc
                    End If
                End If
                If R!cCtaCod = "Total" Or R!cInstitucion = "Total" Then
                   sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
                End If
            
        Else
            
            'Dim lnMoraCalc As Double
            'Dim lnDias As Long
            'Dim lnItfCal As Double
            
            If pdFecIni > pdFecSis And R!nDiasAtraso > 0 Then
                lnDias = DateDiff("d", pdFecSis, pdFecIni)
                lnMoraCalc = Round(lnDias * R!nCapPend * (R!nTasMor / 100), 2)
                lnItfCal = CalculoSinRedondeo(lnMoraCalc)
                If R!cCtaCod = "108013011000002984" Then Stop
            Else
                lnMoraCalc = 0
                lnItfCal = 0
            End If
            
            MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(R!cCtaCod)
            MatGastosCancelacion = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosCancel, gdFecSis, R!cCtaCod, 1, "CA", , , R!nMontoApr, oNegCredito.MatrizMontoCapitalAPagar(MatCalend, gdFecSis), MatCalend(0, 2), , , , , R!nDiasAtraso)
            nMontoGastoCargo = MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("CA", "PA", ""))
            
            sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(PstaNombre(R!sTitular), 25) _
                                    & ImpreFormat(R!cCodModular, 15) & ImpreFormat(Format(R!dFecDesemb, "dd/mm/yyyy"), 12) _
                                    & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!sCuotaPend, 6, 4) & ImpreFormat(R!nCuota + nMontoGastoCargo + lnMoraCalc + lnItfCal, 6) & ImpreFormat(R!nSaldoCuota + nMontoGastoCargo + lnMoraCalc + lnItfCal, 6) _
                                    & ImpreFormat(R!nITFSalCuota + lnItfCal, 6) & ImpreFormat(R!nSaldoCuota + nMontoGastoCargo + lnMoraCalc + R!nITFSalCuota + lnItfCal, 8) & ImpreFormat(Format(R!dFecPago, "dd/mm/yyyy"), 12) & lnSaltoLinDoc
          
        End If
        
        I = I + 1
        RaiseEvent ProgresoBarra(I, "Reporte de Creditos por Institución", Trim(R!cInstitucion) & " - " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES"))
        
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeCreditosXInstitucion = sCadImp
    
    RaiseEvent FinalizaBarra
End Function
Public Function ImprimeCreditosXInstitucionDet(ByVal pMatAgencias As Variant, _
    ByVal pdFecIni As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pnTipoOrden As Integer, ByVal pnIncluyeMora As Integer, _
    Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "", Optional ByVal MatInstitucion As Variant = "", _
    Optional ByVal pnDia As Integer = 0) As String

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
Dim lnTotal As Long
Dim vExcelObj As Excel.Application
Dim vIni As Long
Dim vItem As Long
Dim vNHC As String
Dim lnTotMoraCal As Currency
Dim lnTotITFCal As Currency
Dim lnTotalInst As Long

Dim lnInst As Long
Dim lnMoraCalc As Double

Dim lnMoraCalcInst As Double
Dim lnITFCalcInst As Double

Dim lnDias As Long
Dim lnItfCal As Double
            

    lnTotalInst = UBound(MatInstitucion)
       
    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosXInstitucionDet(pMatAgencias, pnMoneda, pdFecIni, pnTipoOrden, pnIncluyeMora, MatInstitucion, pdFecSis, pnDia)
    Set oDCred = Nothing
    
    Set oGastos = New NGasto
    Set oNegCredito = New NCredito
        
    If R.RecordCount > 0 Then
        lnTotal = R.RecordCount
        
        RaiseEvent IniciaBarra(lnTotal)
        
        vNHC = App.path & "\spooler\" & IIf(pnDia = 30, "REPORTEINSTITUCION", "REPORTEINSTITUCION1") & Format(pdFecIni, "yyyymmdd") & IIf(pnMoneda = 1, "SOLES", "DOLARES") & ".XLS"
        Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
        'vExcelObj.DisplayAlerts = True
        vExcelObj.DisplayAlerts = False
        
        vExcelObj.Workbooks.Add
    
        vExcelObj.Sheets("Hoja1").Select
        vExcelObj.Sheets("Hoja1").Name = Left(R!cInstitucion, 10) & "_1"
        
        vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
        vExcelObj.Range("A1:IV65536").Font.Size = 8
        vExcelObj.Columns("A:IV").Select
        vExcelObj.Selection.VerticalAlignment = 3
        
        vExcelObj.Columns("A").Select
        vExcelObj.Selection.HorizontalAlignment = 1
        vExcelObj.Columns("B:H").Select
        vExcelObj.Selection.HorizontalAlignment = 1
        
        vExcelObj.Range("A1").Select
        vExcelObj.Range("A1").Font.Bold = True
        vExcelObj.Range("A1").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = UCase(Trim(gsNomCmac))
        
        vExcelObj.Range("M1").Select
        vExcelObj.Range("M1").Font.Bold = True
        vExcelObj.Range("M1").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "Informacion al" & pdFecIni
        
        vExcelObj.Range("G3").Select
        vExcelObj.Range("G3").Font.Bold = True
        vExcelObj.Range("G3").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "REPORTE DE CREDITOS PARA INSTITUCIONES CON CONVENIO"
        
        'vExcelObj.Range("A4").Select
        'vExcelObj.Range("A4").Font.Bold = True
        'vExcelObj.Range("A4").HorizontalAlignment = 1
        'vExcelObj.ActiveCell.value = "AGENCIA : " & Trim(R!sAgencia) & "-" & IIf(pnMoneda = "1", "SOLES", "DOLARES")
        
        'vExcelObj.Range("A5").Select
        'vExcelObj.Range("A5").Font.Bold = True
        'vExcelObj.Range("A5").HorizontalAlignment = 1
        'vExcelObj.ActiveCell.value = "INSTITUCION : " & Trim(R!cInstitucion)
        
        vExcelObj.Range("A6").Select
        vExcelObj.Range("A6").Font.Bold = True
        vExcelObj.Range("A6").ColumnWidth = 20
        vExcelObj.ActiveCell.value = "AGENCIA"
        
        vExcelObj.Range("B6").Select
        vExcelObj.Range("B6").Font.Bold = True
        vExcelObj.Range("B6").ColumnWidth = 30
        vExcelObj.ActiveCell.value = "INSTITUCION"
        
        vExcelObj.Range("C6").Select
        vExcelObj.Range("C6").Font.Bold = True
        vExcelObj.Range("C6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "N° CREDITO"
        
        vExcelObj.Range("D6").Select
        vExcelObj.Range("D6").Font.Bold = True
        vExcelObj.Range("D6").ColumnWidth = 30
        vExcelObj.ActiveCell.value = "NOMBRE CLIENTE"
        
        vExcelObj.Range("E6").Select
        vExcelObj.Range("E6").Font.Bold = True
        vExcelObj.Range("E6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "COD.MODULAR"
        
        vExcelObj.Range("F6").Select
        vExcelObj.Range("F6").Font.Bold = True
        vExcelObj.Range("F6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "FEC.DESEMBOLSO"
        
        vExcelObj.Range("G6").Select
        vExcelObj.Range("G6").Font.Bold = True
        vExcelObj.Range("G6").ColumnWidth = 13
        vExcelObj.ActiveCell.value = "SALDO CAPITAL"
                                
        vExcelObj.Range("H6").Select
        vExcelObj.Range("H6").Font.Bold = True
        vExcelObj.Range("H6").ColumnWidth = 13
        vExcelObj.ActiveCell.value = "CUOTA.PENDIENTE"
                                
        vExcelObj.Range("I6").Select
        vExcelObj.Range("I6").Font.Bold = True
        vExcelObj.Range("I6").ColumnWidth = 13
        vExcelObj.ActiveCell.value = "DIAS ATRASO CUOTA"
        
        vExcelObj.Range("J6").Select
        vExcelObj.Range("J6").Font.Bold = True
        vExcelObj.Range("J6").ColumnWidth = 13
        vExcelObj.ActiveCell.value = "CUOTA PACTADA"
        
        vExcelObj.Range("K6").Select
        vExcelObj.Range("K6").Font.Bold = True
        vExcelObj.Range("K6").ColumnWidth = 13
        vExcelObj.ActiveCell.value = "MONTO CUOTA CON MORA"
        
        vExcelObj.Range("L6").Select
        vExcelObj.Range("L6").Font.Bold = True
        vExcelObj.Range("L6").ColumnWidth = 13
        vExcelObj.ActiveCell.value = "SALDO CUOTA FECHA"

        vExcelObj.Range("M6").Select
        vExcelObj.Range("M6").Font.Bold = True
        vExcelObj.Range("M6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "ITF CUOTA"
        
        vExcelObj.Range("N6").Select
        vExcelObj.Range("N6").Font.Bold = True
        vExcelObj.Range("N6").ColumnWidth = 13
        vExcelObj.ActiveCell.value = "TOTAL + ITF"
        
        vExcelObj.Range("O6").Select
        vExcelObj.Range("O6").Font.Bold = True
        vExcelObj.Range("O6").ColumnWidth = 11
        vExcelObj.ActiveCell.value = "VENCIMIENTO"
        
        vExcelObj.Range("P6").Select
        vExcelObj.Range("P6").Font.Bold = True
        vExcelObj.Range("P6").ColumnWidth = 13
        vExcelObj.ActiveCell.value = "MONTO DESCUENTO"
    
    End If
    vIni = 6
    vItem = vIni
    
    Dim vCel  As String
    lnInst = 1 'lnTotalInst = 1
    Do While Not R.EOF
            vItem = vItem + 1
            
            vCel = "A" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = "'" + R!sAgencia
            
            vCel = "B" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = "'" + Trim(R!cInstitucion)
            
                
            vCel = "C" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = "'" + R!cCtaCod

            vCel = "D" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = "'" + PstaNombre(R!sTitular)

            vCel = "E" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = "'" + R!cCodModular

            vCel = "F" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = "'" + Format(R!dFecDesemb, "dd/mm/yyyy")

            vCel = "G" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = Format(R!nSaldoCap, "#0.00")

             vCel = "H" + Trim(Str(vItem))
             vExcelObj.Range(vCel).Select
             vExcelObj.ActiveCell.value = "'" & R!sCuotaPend

             vCel = "I" + Trim(Str(vItem))
             vExcelObj.Range(vCel).Select
             vExcelObj.ActiveCell.value = Format(R!nDiaAtraCuota, "#0")

             vCel = "J" + Trim(Str(vItem))
             vExcelObj.Range(vCel).Select
             vExcelObj.ActiveCell.value = ImpreFormat(R!nCuotaPactada + nMontoGastoCargo, 6)
             
             vCel = "K" + Trim(Str(vItem))
             vExcelObj.Range(vCel).Select
             vExcelObj.ActiveCell.value = ImpreFormat(R!nCuota + nMontoGastoCargo, 6)

             vCel = "L" + Trim(Str(vItem))
             vExcelObj.Range(vCel).Select
             vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + nMontoGastoCargo + R!nMoraCalc, 6)

             vCel = "M" + Trim(Str(vItem))
             vExcelObj.Range(vCel).Select
             vExcelObj.ActiveCell.value = ImpreFormat(R!nITFSalCuota + R!nItfCalc, 6)

             vCel = "N" + Trim(Str(vItem))
             vExcelObj.Range(vCel).Select
             vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + nMontoGastoCargo + R!nMoraCalc + R!nITFSalCuota + R!nItfCalc, 8)

             vCel = "O" + Trim(Str(vItem))
             vExcelObj.Range(vCel).Select
             vExcelObj.ActiveCell.value = ImpreFormat(Format(R!dFecPago, "dd/mm/yyyy"), 12)
        
        I = I + 1
        RaiseEvent ProgresoBarra(I, "Reporte de Creditos por Institución", Trim(R!cInstitucion) & " - " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES"))
        R.MoveNext
    Loop
    If R.RecordCount > 0 Then
        'LSDO
        'INI
        'vExcelObj.Range("A6").Select
        'vExcelObj.Range("A6").SubTotal 4, xlSum, Array(10, 11, 12, 13)
        'FIN
        
        vExcelObj.Range("A1").Select
        vExcelObj.ActiveWorkbook.SaveAs (vNHC)
        vExcelObj.ActiveWorkbook.Close
        vExcelObj.Workbooks.Open (vNHC)
        vExcelObj.Visible = True
    
        Set vExcelObj = Nothing
    End If
    R.Close
    Set R = Nothing
RaiseEvent FinalizaBarra
End Function

Public Function ImprimeCreditosXInstitucionDetAnt(ByVal pMatAgencias As Variant, _
    ByVal pdFecIni As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pnTipoOrden As Integer, ByVal pnIncluyeMora As Integer, _
    Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "", Optional ByVal MatInstitucion As Variant = "", _
    Optional ByVal pnDia As Integer = 0) As String

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
Dim lnTotal As Long
Dim vExcelObj As Excel.Application
Dim vIni As Long
Dim vItem As Long
Dim vNHC As String
Dim lnTotMoraCal As Currency
Dim lnTotITFCal As Currency
Dim lnTotalInst As Long

Dim lnInst As Long
Dim lnMoraCalc As Double

Dim lnMoraCalcInst As Double
Dim lnITFCalcInst As Double

Dim lnDias As Long
Dim lnItfCal As Double
            

    lnTotalInst = UBound(MatInstitucion)
       
    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosXInstitucionDet(pMatAgencias, pnMoneda, pdFecIni, pnTipoOrden, pnIncluyeMora, MatInstitucion, pdFecSis, pnDia)
    Set oDCred = Nothing
    
    Set oGastos = New NGasto
    Set oNegCredito = New NCredito
        
    If R.RecordCount > 0 Then
        lnTotal = R.RecordCount
        
        RaiseEvent IniciaBarra(lnTotal)
        
        vNHC = App.path & "\spooler\" & IIf(pnDia = 30, "REPORTEINSTITUCION", "REPORTEINSTITUCION1") & Format(pdFecIni, "yyyymmdd") & IIf(pnMoneda = 1, "SOLES", "DOLARES") & ".XLS"
        Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
        'vExcelObj.DisplayAlerts = True
        vExcelObj.DisplayAlerts = False
        
        vExcelObj.Workbooks.Add
    
        vExcelObj.Sheets("Hoja1").Select
        vExcelObj.Sheets("Hoja1").Name = Left(R!cInstitucion, 10) & "_1"
        
        vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
        vExcelObj.Range("A1:IV65536").Font.Size = 8
        vExcelObj.Columns("A:IV").Select
        vExcelObj.Selection.VerticalAlignment = 3
        
        vExcelObj.Columns("A").Select
        vExcelObj.Selection.HorizontalAlignment = 1
        vExcelObj.Columns("B:H").Select
        vExcelObj.Selection.HorizontalAlignment = 1
        
        vExcelObj.Range("A1").Select
        vExcelObj.Range("A1").Font.Bold = True
        vExcelObj.Range("A1").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = UCase(Trim(gsNomCmac))
        
        vExcelObj.Range("K1").Select
        vExcelObj.Range("K1").Font.Bold = True
        vExcelObj.Range("K1").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "Informacion al" & pdFecIni
        
        vExcelObj.Range("G3").Select
        vExcelObj.Range("G3").Font.Bold = True
        vExcelObj.Range("G3").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "REPORTE DE CREDITOS PARA INSTITUCIONES CON CONVENIO"
        
        vExcelObj.Range("A4").Select
        vExcelObj.Range("A4").Font.Bold = True
        vExcelObj.Range("A4").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "AGENCIA : " & Trim(R!sAgencia) & "-" & IIf(pnMoneda = "1", "SOLES", "DOLARES")
        
        
        vExcelObj.Range("A5").Select
        vExcelObj.Range("A5").Font.Bold = True
        vExcelObj.Range("A5").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "INSTITUCION : " & Trim(R!cInstitucion)
        
        vExcelObj.Range("A6").Select
        vExcelObj.Range("A6").Font.Bold = True
        vExcelObj.Range("A6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "N° CREDITO"
        
        vExcelObj.Range("B6").Select
        vExcelObj.Range("B6").Font.Bold = True
        vExcelObj.Range("B6").ColumnWidth = 30
        vExcelObj.ActiveCell.value = "Nombre Cliente"
        
        vExcelObj.Range("C6").Select
        vExcelObj.Range("C6").Font.Bold = True
        vExcelObj.Range("C6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Codigo.Modular."
        
        vExcelObj.Range("D6").Select
        vExcelObj.Range("D6").Font.Bold = True
        vExcelObj.Range("D6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "FECHA.DESEMBOLSO"
        
        vExcelObj.Range("E6").Select
        vExcelObj.Range("E6").Font.Bold = True
        vExcelObj.Range("E6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "SALDO CAPITAL"
                                
        vExcelObj.Range("F6").Select
        vExcelObj.Range("F6").Font.Bold = True
        vExcelObj.Range("F6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "CUOTA.PENDIENTE"
                                
        vExcelObj.Range("G6").Select
        vExcelObj.Range("G6").Font.Bold = True
        vExcelObj.Range("G6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "DIAS ATRASO CUOTA"
        
        vExcelObj.Range("H6").Select
        vExcelObj.Range("H6").Font.Bold = True
        vExcelObj.Range("H6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "MONTO CUOTA NORMAL"
        
        vExcelObj.Range("I6").Select
        vExcelObj.Range("I6").Font.Bold = True
        vExcelObj.Range("I6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "SALDO CUOTA FECHA"

        vExcelObj.Range("J6").Select
        vExcelObj.Range("J6").Font.Bold = True
        vExcelObj.Range("J6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "ITF CUOTA"
        
        vExcelObj.Range("K6").Select
        vExcelObj.Range("K6").Font.Bold = True
        vExcelObj.Range("K6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "TOTAL + ITF"
        
        vExcelObj.Range("L6").Select
        vExcelObj.Range("L6").Font.Bold = True
        vExcelObj.Range("L6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "VENCIMIENTO"
        
        vExcelObj.Range("M6").Select
        vExcelObj.Range("M6").Font.Bold = True
        vExcelObj.Range("M6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "MONTO DESCUENTO"
    
    End If
    vIni = 6
    vItem = vIni
    
    Dim vCel  As String
    lnInst = 1 'lnTotalInst = 1
    Do While Not R.EOF
    'FINAL DE INSTITUCION
        vItem = vItem + 1
                 
        If Trim(R!cCtaCod) = "ZTOTAL" Then
            
            vCel = "A" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = 1
            vExcelObj.ActiveCell.value = ImpreFormat("*** SUB TOTAL INSTITUCION ", 42)
            
            vCel = "C" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = 1
            vExcelObj.ActiveCell.value = ImpreFormat("CASOS : " & R!nCasos, 39)
            
            vCel = "D" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = 1
            vExcelObj.ActiveCell.value = ImpreFormat(R!nMontoApr, 10)
        
            vCel = "E" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = 1
            vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCap, 12)
            
            vCel = "H" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = 1
            vExcelObj.ActiveCell.value = ImpreFormat(R!nCuota, 16)
            
            vCel = "I" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + R!nMoraCalc, 6)
            
            vCel = "J" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = ImpreFormat(R!nITFSalCuota + R!nItfCalc, 6)
            
            vCel = "K" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + R!nMoraCalc + R!nITFSalCuota + R!nItfCalc, 8)
            
            lnMoraCalcInst = 0
            lnITFCalcInst = 0
            
            R.MoveNext
            
            vItem = vItem + 1
            
            Dim vNumSheet As String
            If lnInst < lnTotalInst Then
                 vExcelObj.Range("B6").Select
                 vExcelObj.Range("B6").SubTotal 2, xlSum, Array(8, 9), True
                 vExcelObj.Range("B6").SubTotal 2, xlCount, Array(3), False
                 
                'ADICIONAMOS UNA NUEVA HOJA EXCEL
                 lnInst = lnInst + 1
                 vExcelObj.Sheets.Add
                vNumSheet = Trim(Str(lnInst))
                vExcelObj.Sheets("Hoja" + vNumSheet).Select
                vExcelObj.Sheets("Hoja" + vNumSheet).Name = Left(R!cInstitucion, 10) & "_" & vNumSheet
                
                vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
                vExcelObj.Range("A1:IV65536").Font.Size = 8
                vExcelObj.Columns("A:IV").Select
                vExcelObj.Selection.VerticalAlignment = 3
                
                vExcelObj.Columns("A").Select
                vExcelObj.Selection.HorizontalAlignment = 1
                vExcelObj.Columns("B:H").Select
                vExcelObj.Selection.HorizontalAlignment = 1
                
                vExcelObj.Range("A1").Select
                vExcelObj.Range("A1").Font.Bold = True
                vExcelObj.Range("A1").HorizontalAlignment = 1
                vExcelObj.ActiveCell.value = UCase(Trim(gsNomCmac))
                
                vExcelObj.Range("K1").Select
                vExcelObj.Range("K1").Font.Bold = True
                vExcelObj.Range("K1").HorizontalAlignment = 1
                vExcelObj.ActiveCell.value = "Informacion al" & pdFecIni
                
                vExcelObj.Range("G3").Select
                vExcelObj.Range("G3").Font.Bold = True
                vExcelObj.Range("G3").HorizontalAlignment = 1
                vExcelObj.ActiveCell.value = "REPORTE DE CREDITOS PARA INSTITUCIONES CON CONVENIO"
                
                vExcelObj.Range("A4").Select
                vExcelObj.Range("A4").Font.Bold = True
                vExcelObj.Range("A4").HorizontalAlignment = 1
                vExcelObj.ActiveCell.value = "AGENCIA : " & Trim(R!sAgencia) & "-" & IIf(pnMoneda = "1", "SOLES", "DOLARES")
                
                
                vExcelObj.Range("A5").Select
                vExcelObj.Range("A5").Font.Bold = True
                vExcelObj.Range("A5").HorizontalAlignment = 1
                vExcelObj.ActiveCell.value = "INSTITUCION : " & Trim(R!cInstitucion)
                
                vExcelObj.Range("A6").Select
                vExcelObj.Range("A6").Font.Bold = True
                vExcelObj.Range("A6").ColumnWidth = 10
                vExcelObj.ActiveCell.value = "N° CREDITO"
                
                vExcelObj.Range("B6").Select
                vExcelObj.Range("B6").Font.Bold = True
                vExcelObj.Range("B6").ColumnWidth = 30
                vExcelObj.ActiveCell.value = "Nombre Cliente"
                
                vExcelObj.Range("C6").Select
                vExcelObj.Range("C6").Font.Bold = True
                vExcelObj.Range("C6").ColumnWidth = 10
                vExcelObj.ActiveCell.value = "Codigo.Modular."
                
                vExcelObj.Range("D6").Select
                vExcelObj.Range("D6").Font.Bold = True
                vExcelObj.Range("D6").ColumnWidth = 10
                vExcelObj.ActiveCell.value = "FECHA.DESEMBOLSO"
                
                vExcelObj.Range("E6").Select
                vExcelObj.Range("E6").Font.Bold = True
                vExcelObj.Range("E6").ColumnWidth = 10
                vExcelObj.ActiveCell.value = "SALDO CAPITAL"
                                        
                vExcelObj.Range("F6").Select
                vExcelObj.Range("F6").Font.Bold = True
                vExcelObj.Range("F6").ColumnWidth = 10
                vExcelObj.ActiveCell.value = "CUOTA.PENDIENTE"
                                        
                vExcelObj.Range("G6").Select
                vExcelObj.Range("G6").Font.Bold = True
                vExcelObj.Range("G6").ColumnWidth = 10
                vExcelObj.ActiveCell.value = "DIAS ATRASO CUOTA"
                
                vExcelObj.Range("H6").Select
                vExcelObj.Range("H6").Font.Bold = True
                vExcelObj.Range("H6").ColumnWidth = 10
                vExcelObj.ActiveCell.value = "MONTO CUOTA NORMAL"
                
                vExcelObj.Range("I6").Select
                vExcelObj.Range("I6").Font.Bold = True
                vExcelObj.Range("I6").ColumnWidth = 10
                vExcelObj.ActiveCell.value = "SALDO CUOTA FECHA"
        
                vExcelObj.Range("J6").Select
                vExcelObj.Range("J6").Font.Bold = True
                vExcelObj.Range("J6").ColumnWidth = 10
                vExcelObj.ActiveCell.value = "ITF CUOTA"
                
                vExcelObj.Range("K6").Select
                vExcelObj.Range("K6").Font.Bold = True
                vExcelObj.Range("K6").ColumnWidth = 10
                vExcelObj.ActiveCell.value = "TOTAL + ITF"
                
                vExcelObj.Range("L6").Select
                vExcelObj.Range("L6").Font.Bold = True
                vExcelObj.Range("L6").ColumnWidth = 10
                vExcelObj.ActiveCell.value = "VENCIMIENTO"
                
                vExcelObj.Range("M6").Select
                vExcelObj.Range("M6").Font.Bold = True
                vExcelObj.Range("M6").ColumnWidth = 10
                vExcelObj.ActiveCell.value = "MONTO DESCUENTO"
                
                
                vItem = vIni
            End If
    
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE AGENCIA
            If R!cInstitucion = "ZTOTAL" Then
                vCel = "A" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.Range(vCel).Font.Bold = True
                vExcelObj.Range(vCel).HorizontalAlignment = 1
                vExcelObj.ActiveCell.value = ImpreFormat("*** SUB TOTAL AGENCIA ", 42)

                vCel = "C" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.Range(vCel).Font.Bold = True
                vExcelObj.Range(vCel).HorizontalAlignment = 1
                vExcelObj.ActiveCell.value = ImpreFormat("CASOS : " & R!nCasos, 39)

                vCel = "D" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.Range(vCel).Font.Bold = True
                vExcelObj.Range(vCel).HorizontalAlignment = 1
                vExcelObj.ActiveCell.value = ImpreFormat(R!nMontoApr, 10)

                vCel = "E" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.Range(vCel).Font.Bold = True
                vExcelObj.Range(vCel).HorizontalAlignment = 1
                vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCap, 12)

                vCel = "H" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.Range(vCel).Font.Bold = True
                vExcelObj.Range(vCel).HorizontalAlignment = 1
                vExcelObj.ActiveCell.value = ImpreFormat(R!nCuota, 16)

                vCel = "I" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.Range(vCel).Font.Bold = True
                vExcelObj.Range(vCel).HorizontalAlignment = 1
                vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + R!nMoraCalc, 16)

                vCel = "K" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.Range(vCel).Font.Bold = True
                vExcelObj.Range(vCel).HorizontalAlignment = 1
                vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + R!nMoraCalc + R!nITFSalCuota + R!nItfCalc, 8)

                R.MoveNext
                vItem = vItem + 1
                If R.EOF Then
                    Exit Do
                End If

                'TOTAL FINAL
                If R!sAgencia = "ZTOTAL" Then

                    vCel = "A" + Trim(Str(vItem))
                    vExcelObj.Range(vCel).Select
                    vExcelObj.Range(vCel).Font.Bold = True
                    vExcelObj.Range(vCel).HorizontalAlignment = 1
                    vExcelObj.ActiveCell.value = ImpreFormat("*** TOTAL ", 42)

                    vCel = "C" + Trim(Str(vItem))
                    vExcelObj.Range(vCel).Select
                    vExcelObj.Range(vCel).Font.Bold = True
                    vExcelObj.Range(vCel).HorizontalAlignment = 1
                    vExcelObj.ActiveCell.value = ImpreFormat("CASOS : " & R!nCasos, 39)

                    vCel = "D" + Trim(Str(vItem))
                    vExcelObj.Range(vCel).Select
                    vExcelObj.Range(vCel).Font.Bold = True
                    vExcelObj.Range(vCel).HorizontalAlignment = 1
                    vExcelObj.ActiveCell.value = ""

                    vCel = "E" + Trim(Str(vItem))
                    vExcelObj.Range(vCel).Select
                    vExcelObj.Range(vCel).Font.Bold = True
                    vExcelObj.Range(vCel).HorizontalAlignment = 1
                    vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCap, 12)

                    vCel = "H" + Trim(Str(vItem))
                    vExcelObj.Range(vCel).Select
                    vExcelObj.Range(vCel).Font.Bold = True
                    vExcelObj.Range(vCel).HorizontalAlignment = 1
                    vExcelObj.ActiveCell.value = ImpreFormat(R!nCuota, 16)

                    vCel = "I" + Trim(Str(vItem))
                    vExcelObj.Range(vCel).Select
                    vExcelObj.Range(vCel).Font.Bold = True
                    vExcelObj.Range(vCel).HorizontalAlignment = 1
                    vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + R!nMoraCalc, 16)

                    vCel = "J" + Trim(Str(vItem))
                    vExcelObj.Range(vCel).Select
                    vExcelObj.Range(vCel).Font.Bold = True
                    vExcelObj.Range(vCel).HorizontalAlignment = 1
                    vExcelObj.ActiveCell.value = ImpreFormat(R!nITFSalCuota + R!nItfCalc, 16)

                    vCel = "K" + Trim(Str(vItem))
                    vExcelObj.Range(vCel).Select
                    vExcelObj.Range(vCel).Font.Bold = True
                    vExcelObj.Range(vCel).HorizontalAlignment = 1
                    vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + R!nMoraCalc + R!nITFSalCuota + R!nItfCalc, 8) 'ImpreFormat(R!nSaldoCuota, 16)

                    Exit Do
                Else
                    nPuntPag = 0
                    vItem = vItem + 1
                    vCel = "A" + Trim(Str(vItem + 1))
                    vExcelObj.Range(vCel).Select
                    vExcelObj.Range(vCel).Font.Bold = True
                    vExcelObj.Range(vCel).HorizontalAlignment = 1
                    vExcelObj.ActiveCell.value = "INSTITUCION : " & Trim(R!cInstitucion)
                End If
            Else
                vItem = vItem + 1

                vCel = "A" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = "'" + R!cCtaCod

                vCel = "B" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = "'" + PstaNombre(R!sTitular)

                vCel = "C" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = "'" + R!cCodModular

                vCel = "D" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = "'" + Format(R!dFecDesemb, "dd/mm/yyyy")

                vCel = "E" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = Format(R!nSaldoCap, "#0.00")

                 vCel = "F" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = "'" & R!sCuotaPend

                 vCel = "G" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = Format(R!nDiaAtraCuota, "#0")

                 vCel = "H" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(R!nCuota + nMontoGastoCargo, 6)

                 vCel = "I" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + nMontoGastoCargo + R!nMoraCalc, 6)

                 vCel = "J" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(R!nITFSalCuota + R!nItfCalc, 6)

                 vCel = "K" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + nMontoGastoCargo + R!nMoraCalc + R!nITFSalCuota + R!nItfCalc, 8)

                 vCel = "L" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(Format(R!dFecPago, "dd/mm/yyyy"), 12)
                
            End If
        Else
            
            'MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(R!Cctacod)
            'MatGastosCancelacion = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosCancel, gdFecSis, R!Cctacod, 1, "CA", , , R!nMontoApr, oNegCredito.MatrizMontoCapitalAPagar(MatCalend, gdFecSis), MatCalend(0, 2), , , , , R!nDiasAtraso)
            'nMontoGastoCargo = MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("CA", "PA", ""))
            
            If R!cTotCuota = "" Then
            
                vCel = "A" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = "'" + R!cCtaCod
        
                vCel = "B" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = "'" + PstaNombre(R!sTitular)
        
                vCel = "C" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = "'" + R!cCodModular
        
                vCel = "D" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = "'" + Format(R!dFecDesemb, "dd/mm/yyyy")
        
                vCel = "E" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = Format(R!nSaldoCap, "#0.00")
        
                 vCel = "F" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = "'" & R!sCuotaPend
            
                 vCel = "G" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = Format(R!nDiaAtraCuota, "#0")
                 
                 vCel = "H" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(R!nCuota + nMontoGastoCargo, 6)
             
             
                 vCel = "I" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + nMontoGastoCargo + R!nMoraCalc, 6)
                 
                 vCel = "J" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(R!nITFSalCuota + R!nItfCalc, 6)
                 
                 vCel = "K" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + nMontoGastoCargo + R!nMoraCalc + R!nITFSalCuota + R!nItfCalc, 8)
                 
                 vCel = "L" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(Format(R!dFecPago, "dd/mm/yyyy"), 12)
            Else
                
                vCel = "B" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Font.Bold = True
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = "TOTAL DEUDA CLIENTE"

                 vCel = "F" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Font.Bold = True
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = "'" & R!sCuotaPend

                 vCel = "G" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Font.Bold = True
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = Format(R!nDiaAtraCuota, "#0")

                 vCel = "I" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Font.Bold = True
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + nMontoGastoCargo + R!nMoraCalc, 6)

                 vCel = "J" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Font.Bold = True
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(R!nITFSalCuota + R!nItfCalc, 6)

                 vCel = "K" + Trim(Str(vItem))
                 vExcelObj.Range(vCel).Font.Bold = True
                 vExcelObj.Range(vCel).Select
                 vExcelObj.ActiveCell.value = ImpreFormat(R!nSaldoCuota + nMontoGastoCargo + R!nMoraCalc + R!nITFSalCuota + R!nItfCalc, 8)

            End If
        End If
        I = I + 1
        RaiseEvent ProgresoBarra(I, "Reporte de Creditos por Institución", Trim(R!cInstitucion) & " - " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES"))
        R.MoveNext
    Loop
    If R.RecordCount > 0 Then
    
        'If Dir(vNHC) <> "" Then
        '   If MsgBox("Archivo Ya Existe ...  Desea Reemplazarlo ??", vbQuestion + vbYesNo + vbDefaultButton1, " Mensaje del Sistema ...") = vbNo Then
        '      Exit Function
        '   End If
        'End If
        vExcelObj.Range("A1").Select
        vExcelObj.ActiveWorkbook.SaveAs (vNHC)
        vExcelObj.ActiveWorkbook.Close
        vExcelObj.Workbooks.Open (vNHC)
        vExcelObj.Visible = True
    
        Set vExcelObj = Nothing
    End If
    R.Close
    Set R = Nothing
RaiseEvent FinalizaBarra
End Function


Function ListaCreditosComite(ByVal pdFecSis As Date, ByVal psCodAgen As String) As String
   Dim rs As ADODB.Recordset
   Dim rs_Garant As ADODB.Recordset
   Dim odCredDoc As DCredDoc
   Dim m_Excel As Excel.Application
   Dim oLibroExcel As Excel.Workbook
   Dim oHojaExcel As Excel.Worksheet
   Dim oCelda As Excel.Range
   Dim nCont As Integer
   
   Dim nFila As Integer
   Dim bNewFila As Boolean
   
   bNewFila = False
   
   'IniciaBarra
   
   Set odCredDoc = New DCredDoc
   Set rs = odCredDoc.RecuperaListaCreditosComite(pdFecSis, psCodAgen)
   Set odCredDoc = Nothing
   
   Set m_Excel = New Excel.Application
   m_Excel.Visible = True
   
   
   Set oLibroExcel = m_Excel.Workbooks.Add
   Set oHojaExcel = oLibroExcel.Worksheets(1)
   oHojaExcel.Visible = xlSheetVisible
   
   
   oHojaExcel.Activate
   
   oHojaExcel.PageSetup.Zoom = 80
   oHojaExcel.PageSetup.Orientation = xlLandscape
   m_Excel.Range("A1:R1000").Font.Size = 9
   
  m_Excel.Selection.NumberFormat = "#,##0.00"
   
   'creando el encabezado del Informe
   oHojaExcel.Range("A2:D2").Merge
   oHojaExcel.Range("A2.D2").value = "ACTA DE COMITE DEL CREDITOS 2005"
   oHojaExcel.Range("A2:D2").Font.Italic = True
   oHojaExcel.Range("A2:D2").Font.Size = 13
   
   
   oHojaExcel.Cells(3, 1) = "Siendo las " & Time & " Los miembros del Comite de Creditos " & _
                             " en su Conjunto se reunierón para aprobar los siguientes Creditos"
                             
   nFila = 3
  'Armando las Columnas
   nFila = nFila + 2
   
   With oHojaExcel
    .Range("A1").ColumnWidth = 5 '1
    .Range("B1").ColumnWidth = 5 '1
    .Range("C1").ColumnWidth = 10 '2
    .Range("D1").ColumnWidth = 10 '3
    .Range("E1").ColumnWidth = 10 '4
    .Range("F1").ColumnWidth = 10 '5
    .Range("G1").ColumnWidth = 10 '6
    .Range("H1").ColumnWidth = 10 '7
    .Range("I1").ColumnWidth = 10 '8
    .Range("J1").ColumnWidth = 10 '9
    .Range("K1").ColumnWidth = 10 '10
    .Range("B1").EntireColumn.NumberFormat = "@"
    .Range("J1").EntireColumn.NumberFormat = "0.00"
    .Range("I1").EntireColumn.NumberFormat = "0.00"
   End With
       
          
   With oHojaExcel
    .Cells(nFila, 1) = "Nro" 'A
    .Cells(nFila, 1).Font.Bold = True
    .Cells(nFila, 2) = "Nro de Credito" 'A
    .Cells(nFila, 2).Font.Bold = True
    .Cells(nFila, 3) = "Nombre del Cliente" 'A
    .Cells(nFila, 3).Font.Bold = True
    .Cells(nFila, 4) = "Tipo de Crédito" 'B
    .Cells(nFila, 4).Font.Bold = True
    .Cells(nFila, 5) = "Condición" 'C
    .Cells(nFila, 5).Font.Bold = True
    .Cells(nFila, 6) = "Monto Aprobado" 'G
    .Cells(nFila, 6).Font.Bold = True
    .Cells(nFila, 7) = "Cuotas Aprobadas" 'H
    .Cells(nFila, 7).Font.Bold = True
    .Cells(nFila, 8) = "Garantia" 'K
    .Cells(nFila, 8).Font.Bold = True
    .Cells(nFila, 9) = "Monto Garantia" 'K
    .Cells(nFila, 9).Font.Bold = True
    .Cells(nFila, 10) = "Disponible Garant" 'J
    .Cells(nFila, 10).Font.Bold = True
    .Cells(nFila, 11) = "Cod.Ana" 'J
    .Cells(nFila, 11).Font.Bold = True
   End With
   
   'Set odCredDoc = New DCredDoc
   'Set rs = odCredDoc.RecuperaListaCreditosComite(pdFecSis)
   'Set odCredDoc = Nothing
   
   nCont = nCont + 1
   nFila = nFila + 1
   Do Until rs.EOF
     With oHojaExcel
        '.Cells.EntireColumn.AutoFit
        .Cells(nFila, 1) = nCont
        .Cells(nFila, 2) = Trim(rs!cCtaCod)
        .Cells(nFila, 3) = Trim(rs!Titular)
        .Cells(nFila, 4) = rs!TipoCredito
        .Cells(nFila, 5) = rs!Condicion
        '.Cells(nFila, 4) = IIf(Mid(rs!cCtaCod, 9, 1) = "1", "S/.", "$.") & Format(CDbl(rs!MontoPropuesto), "#,##0.00")
        '.Cells(nFila, 5) = rs!CuotasSugeridas
'        If Val(rs!PlazoPropuesto) = 0 Then
'            .Cells(nFila, 6) = "D- " & rs!FechaFijaSugerido
'        Else
'            .Cells(nFila, 6) = "C-" & rs!PlazoPropuesto
'        End If
        .Cells(nFila, 6) = IIf(Mid(rs!cCtaCod, 9, 1) = "1", "S/.", "$.") & Format(CDbl(rs!MontoAprobado), "#,#0.00")
        .Cells(nFila, 7) = rs!nCuotasAprobadas
'        If Val(rs!PlazoAprobado) = 0 Then
'            .Cells(nFila, 9) = "D- " & rs!FechaFijaAprobado
'        Else
'            .Cells(nFila, 9) = "C-" & rs!PlazoAprobado
'        End If
        .Cells(nFila, 11) = rs!Analista
'        .Cells(nFila, 9) = ListaDegarantiasCredito(rs!cCtaCod) & Format(IIf(IsNull(rs!MontoGarantia), 0, rs!MontoGarantia), "#,#0.00")
        Set rs_Garant = ListaDegarantiasCredito(rs!cCtaCod)
        Do Until rs_Garant.EOF
            '.Cells(nFila, 12) = .Cells(nFila, 10) & "" & rs_Garant!cConsDescripcion
            .Cells(nFila, 8) = rs_Garant!cConsDescripcion
            .Cells(nFila, 9) = rs_Garant!nGravado
            .Cells(nFila, 10) = rs_Garant!nDisponible
            nFila = nFila + 1
            bNewFila = True
           rs_Garant.MoveNext
        Loop
        Set rs_Garant = Nothing
     End With
     If bNewFila = False Then
        nFila = nFila + 1
     End If
     bNewFila = False
     nCont = nCont + 1
     rs.MoveNext
   '  ProgresoBarra
   Loop
   Set rs = Nothing
   oLibroExcel.PrintPreview
   
End Function
Function ListaDegarantiasCredito(ByVal psCtaCod As String) As Recordset
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Dim oConec As DConecta
    
    
    Set oConec = New DConecta
    sSQL = "Select CN.cConsDescripcion,CG.nGravado,G.nPorGravar-G.nGravament as nDisponible"
    sSQL = sSQL & " From ColocGarantia CG"
    sSQL = sSQL & " Inner Join Garantias G on G.cNumGarant=CG.cNumGarant"
    sSQL = sSQL & " Inner Join Constante CN on CN.nConsCod=1027 and CN.nConsValor<>1027 and CN.nConsValor=G.nTpoGarantia"
    sSQL = sSQL & " Where CG.cCtaCod='" & psCtaCod & "'"
'

'    sSQL = "Select top 1 G.nMoneda"
'    sSQL = sSQL & " From ColocGarantia CG"
'    sSQL = sSQL & " Inner Join Garantias G on G.cNumGarant=CG.cNumGarant"
'    sSQL = sSQL & " Inner Join Constante CN on CN.nConsCod=1027 and CN.nConsValor<>1027 and CN.nConsValor=G.nTpoGarantia"
'    sSQL = sSQL & " Where CG.cCtaCod='" & psCtacod & "'"
    oConec.AbreConexion
    Set ListaDegarantiasCredito = oConec.CargaRecordSet(sSQL)
'    If Not rs.EOF And Not rs.BOF Then
'        If rs!nMoneda = "1" Then
'           ListaDegarantiasCredito = "S/."
'        Else
'           ListaDegarantiasCredito = "$."
'        End If
    'End If
   ' Set rs = Nothing
    oConec.CierraConexion
    Set oConec = Nothing
End Function


Function CalculoSinRedondeo(ByVal pnMonto As Double) As Double
    Dim sCadena As String
    Dim intpos  As Integer
    Dim nEntera As Integer
    Dim nDecimal As Integer
    Dim lnValor As Double
        
        fgITFParametros
        
        lnValor = pnMonto * gnITFPorcent
        lnValor = CortaDosITF(lnValor)
        lnValor = Format(lnValor, "#0.00")
        CalculoSinRedondeo = lnValor
       
End Function


'
'Public Function ImprimeCreditosXInstitucion(ByVal pMatAgencias As Variant, _
'    ByVal pdFecIni As Date, ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
'    ByVal psNomAge As String, ByVal pnTipoOrden As Integer, ByVal pnIncluyeMora As Integer, _
'    Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "") As String
'
'Dim oDCred As DCredDoc
'Dim R As ADODB.Recordset
'Dim sCadImp As String
'Dim I As Integer
'Dim sAgencia As String
'Dim oGastos As NGasto
'Dim oNegCredito As NCredito
'Dim MatCalend As Variant
'Dim MatGastosCancelacion As Variant
'Dim nMontoGastoCargo As Double
'Dim nNumGastosCancel As Integer
'
'    Set oDCred = New DCredDoc
'    Set R = oDCred.RecuperaCreditosXInstitucion(pMatAgencias, pnMoneda, pdFecIni, pnTipoOrden, pnIncluyeMora)
'    Set oDCred = Nothing
'
'    Set oGastos = New NGasto
'    Set oNegCredito = New NCredito
'
'    If R.RecordCount > 0 Then
'        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "LISTADO PARA DESCUENTO POR PLANILLA AL : " & Format(pdFecIni, "dd/mm/yyyy"), psNomCmac, , True, psCodRepo)
'        sCadImp = sCadImp & lnSaltoLinDoc
'        sCadImp = sCadImp & lnSaltoLinDoc
'        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
'        sCadImp = sCadImp & lnSaltoLinDoc
'        sCadImp = sCadImp & lnSaltoLinDoc
'        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
'        sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
'        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
'        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 25) & ImpreFormat("COD. MODULAR", 20) & _
'            ImpreFormat("DESEMBOLSO", 15) & ImpreFormat("APROBADO", 12) & ImpreFormat("SALDO CAP", 10) & ImpreFormat("C. PEND", 7) & _
'            ImpreFormat("CUOTA", 8) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
'        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
'    End If
'
'    Do While Not R.EOF
'    'FINAL DE INSTITUCION
'        If Trim(R!cCtacod) = "TOTAL" Then
'            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
'            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL INSTITUCION ", 42) & ImpreFormat("CASOS : " & R!nCasos, 39) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuota, 16) & lnSaltoLinDoc
'            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
'            sCadImp = sCadImp & lnSaltoLinDoc
'            R.MoveNext
'            If R.EOF Then
'                Exit Do
'            End If
'
'            'FINAL DE AGENCIA
'            If R!cInstitucion = "TOTAL" Then
'                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
'                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 42) & ImpreFormat("CASOS : " & R!nCasos, 39) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuota, 16) & lnSaltoLinDoc
'                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
'                sCadImp = sCadImp & lnSaltoLinDoc
'                R.MoveNext
'                If R.EOF Then
'                    Exit Do
'                End If
'                'TOTAL FINAL
'                If R!sAgencia = "TOTAL" Then
'                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 42) & ImpreFormat("CASOS : " & R!nCasos, 39) & ImpreFormat(R!nMontoApr, 10) & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!nCuota, 16) & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
'                    sCadImp = sCadImp & lnSaltoLinDoc
'                    Exit Do
'                Else
'                    nPuntPag = 0
'                    sCadImp = sCadImp & Chr$(12)
'                    sCadImp = sCadImp & lnSaltoLinDoc
'                    sCadImp = sCadImp & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & lnSaltoLinDoc
'                    sCadImp = sCadImp & lnSaltoLinDoc
'                    sCadImp = sCadImp & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("COD. MODULAR", 20) & _
'                        ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("APROBADO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("C. PEND", 5) & _
'                        ImpreFormat("CUOTA", 12) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
'                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
'                End If
'
'            Else
'                sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
'                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
'                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 34) & ImpreFormat("COD. MODULAR", 20) & _
'                ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("APROBADO", 12) & ImpreFormat("SALDO CAP", 12) & ImpreFormat("C. PEND", 5) & _
'                        ImpreFormat("CUOTA", 12) & ImpreFormat("FEC. PAGO", 12) & lnSaltoLinDoc
'                        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
'            End If
'        End If
'
'        If R!cCtacod <> "TOTAL" Then
'            MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(R!cCtacod)
'            MatGastosCancelacion = oGastos.GeneraCalendarioGastos(Array(0), Array(0), nNumGastosCancel, gdFecSis, R!cCtacod, 1, "CA", , , R!nMontoApr, oNegCredito.MatrizMontoCapitalAPagar(MatCalend, gdFecSis), MatCalend(0, 2), , , , , R!nDiasAtraso)
'            nMontoGastoCargo = MontoTotalGastosGenerado(MatGastosCancelacion, nNumGastosCancel, Array("CA", "PA", ""))
'
'            sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtacod, 20) & ImpreFormat(PstaNombre(R!sTitular), 25) & _
'                                        ImpreFormat(R!cCodModular, 20) & ImpreFormat(Format(R!dFecDesemb, "dd/mm/yyyy"), 12) & ImpreFormat(R!nMontoApr, 10) _
'                                        & ImpreFormat(R!nSaldoCap, 12) & ImpreFormat(R!sCuotaPend, 6, 4) & ImpreFormat(R!nCuota + nMontoGastoCargo, 6) & ImpreFormat(Format(R!dFecPago, "dd/mm/yyyy"), 12) & lnSaltoLinDoc
'        End If
'
'        R.MoveNext
'    Loop
'    R.Close
'    Set R = Nothing
'    ImprimeCreditosXInstitucion = sCadImp
'
'End Function

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
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 27) & ImpreFormat("MONTO", 11) & ImpreFormat("SALDO", 9) & ImpreFormat("SALDO DE", 10, 0) & ImpreFormat("DIAS", 5, 3) & ImpreFormat("SALDO", 7) & ImpreFormat("FECHA DE", 10) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat(" ", 47) & ImpreFormat("DESEMBOLSADO", 11) & ImpreFormat("CAPITAL", 10) & ImpreFormat("CUOTA", 10) & ImpreFormat("MORA", 6) & ImpreFormat("MORA", 7) & ImpreFormat("VENCIMIENTO", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE INSTITUCION
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL INSTITUCION ", 26)
            sCadImp = sCadImp & ImpreFormat("CASOS : " & R!nCasos, 19) & ImpreFormat(Format(R!nMontoDesemb, "#0.00"), 10)
            sCadImp = sCadImp & ImpreFormat(Format(R!nSaldoCap, "#0.00"), 10) & ImpreFormat(Format(R!nSaldoCuota, "#0.00"), 10)
            sCadImp = sCadImp & ImpreFormat(R!nSaldoMora, 13) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE AGENCIA
            If R!cInstitucion = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 26)
                sCadImp = sCadImp & ImpreFormat("CASOS : " & R!nCasos, 19) & ImpreFormat(Format(R!nMontoDesemb, "#0.00"), 10)
                sCadImp = sCadImp & ImpreFormat(Format(R!nSaldoCap, "#0.00"), 10) & ImpreFormat(Format(R!nSaldoCuota, "#0.00"), 10)
                sCadImp = sCadImp & ImpreFormat(R!nSaldoMora, 13) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'TOTAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 26)
                    sCadImp = sCadImp & ImpreFormat("CASOS : " & R!nCasos, 19) & ImpreFormat(Format(R!nMontoDesemb, "#0.00"), 10)
                    sCadImp = sCadImp & ImpreFormat(Format(R!nSaldoCap, "#0.00"), 10) & ImpreFormat(Format(R!nSaldoCuota, "#0.00"), 10)
                    sCadImp = sCadImp & ImpreFormat(R!nSaldoMora, 13) & lnSaltoLinDoc
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
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 27) & ImpreFormat("MONTO", 11) & ImpreFormat("SALDO", 9) & ImpreFormat("SALDO DE", 10, 0) & ImpreFormat("DIAS", 5, 3) & ImpreFormat("SALDO", 7) & ImpreFormat("FECHA DE", 10) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat(" ", 47) & ImpreFormat("DESEMBOLSADO", 11) & ImpreFormat("CAPITAL", 10) & ImpreFormat("CUOTA", 10) & ImpreFormat("MORA", 6) & ImpreFormat("MORA", 7) & ImpreFormat("VENCIMIENTO", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
        
            Else
                sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 31) & ImpreFormat("MONTO", 13) & ImpreFormat("SALDO", 9) & ImpreFormat("SALDO DE", 10, 0) & ImpreFormat("DIAS", 5, 3) & ImpreFormat("SALDO", 7) & ImpreFormat("FECHA DE", 10) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat(" ", 51) & ImpreFormat("DESEMBOLSADO", 13) & ImpreFormat("CAPITAL", 10) & ImpreFormat("CUOTA", 10) & ImpreFormat("MORA", 6) & ImpreFormat("MORA", 7) & ImpreFormat("VENCIMIENTO", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 25) & _
                                    ImpreFormat(Format(R!nMontoDesemb, "#0.00"), 10) & ImpreFormat(Format(R!nSaldoCap, "#0.00"), 10) _
                                    & ImpreFormat(Format(R!nSaldoCuota, "#0.00"), 9) & ImpreFormat(Trim(Str(R!nDiasAtraso)), 4, 6) _
                                    & ImpreFormat(Format(R!nSaldoMora, "#0.00"), 6) & ImpreFormat(Format(R!dFecVencimiento, "dd/mm/yyyy"), 12) & lnSaltoLinDoc
                                    
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
            'R.MoveNext
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
        'nTotalRep = nTotalRep + R!nMonto
        'nTotal = nTotal + R!nMonto
        
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
            sCadImp = sCadImp & ImpreFormat("SUB TOTAL : ", 21) & " CASOS : " & ImpreFormat(nContCasos, 5, 0) & ImpreFormat(nSubTotal, 41)
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            nTotal = nTotal + nSubTotal
            nContTotalCasos = nContTotalCasos + nContCasos
            nContCasos = 0
            nSubTotal = 0
            'R.MoveNext
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
            'R.MoveNext
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
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!cpersnombre, 30) & ImpreFormat(IIf(IsNull(R!cOperacion), "", R!cOperacion), 30) & ImpreFormat(IIf(IsNull(R!cNroDoc), "", R!cNroDoc), 12) _
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
Dim nNumCab As Integer
Dim sCadImpTmp As String
Dim nNumLinTmp As Integer
    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaInteresesEnSuspenso(pMatAgencias, pdFecIni, pnMoneda, pMatProd)
    Set oDCred = Nothing
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "INTERESES EN SUSPENSO DE CREDITOS REFINANCIADOS", psNomCmac, , True)
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & Chr$(10)
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & Chr$(10)
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & Chr$(10)
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
        sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
        
    End If
    nNumCab = 15
    sCadImpTmp = ""
    nNumLinTmp = 0
    Do While Not R.EOF
        'FINAL DE FUENTE DE FINANCIAMIENTO
        If Trim(R!cCtaCod) = "TOTAL" Then
            nNumCab = nNumCab + 4
            nNumLinTmp = nNumLinTmp + 4
            sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
            sCadImpTmp = sCadImpTmp & lsTab & ImpreFormat("*** SUB TOTAL FINANCIAMIENTO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & Chr$(10)
            sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
            sCadImpTmp = sCadImpTmp & Chr$(10)
            
            If nNumCab <= 55 Then
                sCadImp = sCadImp & sCadImpTmp
                sCadImpTmp = ""
                nNumLinTmp = 0
            End If
            
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE PLAZO
            If Trim(R!sFuenteFinan) = "TOTAL" Then
                nNumCab = nNumCab + 4
                nNumLinTmp = nNumLinTmp + 4
                
                sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                sCadImpTmp = sCadImpTmp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & Chr$(10)
                sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                sCadImpTmp = sCadImpTmp & Chr$(10)
                
                If nNumCab <= 55 Then
                    sCadImp = sCadImp & sCadImpTmp
                    sCadImpTmp = ""
                    nNumLinTmp = 0
                End If
                
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If Trim(R!sPlazo) = "TOTAL" Then
                    nNumCab = nNumCab + 4
                    nNumLinTmp = nNumLinTmp + 4
                    
                    sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                    sCadImpTmp = sCadImpTmp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & Chr$(10)
                    sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                    sCadImpTmp = sCadImpTmp & Chr$(10)
                    
                    If nNumCab <= 55 Then
                        sCadImp = sCadImp & sCadImpTmp
                        sCadImpTmp = ""
                        nNumLinTmp = 0
                    End If
                
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL DE AGENCIA
                    If R!sProducto = "TOTAL" Then
                        nNumCab = nNumCab + 4
                        nNumLinTmp = nNumLinTmp + 4
                        
                        sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                        sCadImpTmp = sCadImpTmp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & Chr$(10)
                        sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                        sCadImpTmp = sCadImpTmp & Chr$(10)
                        
                        If nNumCab <= 55 Then
                            sCadImp = sCadImp & sCadImpTmp
                            sCadImpTmp = ""
                            nNumLinTmp = 0
                        End If
                        
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'FINAL
                        If R!sAgencia = "TOTAL" Then
                            nNumCab = nNumCab + 4
                            nNumLinTmp = nNumLinTmp + 4
                            
                            sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                            sCadImpTmp = sCadImpTmp & lsTab & ImpreFormat("*** TOTAL ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nGastos + R!nInteresMor + R!nInteresComp, 12) & Chr$(10)
                            sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                            sCadImpTmp = sCadImpTmp & Chr$(10)
                            
                            If nNumCab <= 55 Then
                                sCadImp = sCadImp & sCadImpTmp
                                sCadImpTmp = ""
                                nNumLinTmp = 0
                            End If
                            Exit Do
                        Else
                            nNumCab = nNumCab + 12
                            nNumLinTmp = nNumLinTmp + 12
                            
                            nPuntPag = 0
                            sCadImpTmp = sCadImpTmp & Chr$(10)
                            sCadImpTmp = sCadImpTmp & Chr$(10)
                            sCadImpTmp = sCadImpTmp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & Chr$(10)
                            sCadImpTmp = sCadImpTmp & Chr$(10)
                            sCadImpTmp = sCadImpTmp & Chr$(10)
                            sCadImpTmp = sCadImpTmp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & Chr$(10)
                            sCadImpTmp = sCadImpTmp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & Chr$(10)
                            sCadImpTmp = sCadImpTmp & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
                            sCadImpTmp = sCadImpTmp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                            sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                            sCadImpTmp = sCadImpTmp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & Chr$(10)
                            sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                            
                            If nNumCab <= 55 Then
                                sCadImp = sCadImp & sCadImpTmp
                                sCadImpTmp = ""
                                nNumLinTmp = 0
                            End If
                        End If
                    Else
                        nNumCab = nNumCab + 6
                        nNumLinTmp = nNumLinTmp + 6
                        
                        sCadImpTmp = sCadImpTmp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & Chr$(10)
                        sCadImpTmp = sCadImpTmp & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
                        sCadImpTmp = sCadImpTmp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                        sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                        sCadImpTmp = sCadImpTmp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & Chr$(10)
                        sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                        
                        If nNumCab <= 55 Then
                            sCadImp = sCadImp & sCadImpTmp
                            sCadImpTmp = ""
                            nNumLinTmp = 0
                        End If
                    End If
                Else
                    nNumCab = nNumCab + 5
                    nNumLinTmp = nNumLinTmp + 5
                    
                    sCadImpTmp = sCadImpTmp & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
                    sCadImpTmp = sCadImpTmp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                    sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                    sCadImpTmp = sCadImpTmp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & Chr$(10)
                    sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                    If nNumCab <= 55 Then
                        sCadImp = sCadImp & sCadImpTmp
                        sCadImpTmp = ""
                        nNumLinTmp = 0
                    End If
                End If
            Else
                nNumCab = nNumCab + 4
                nNumLinTmp = nNumLinTmp + 4
                
                sCadImpTmp = sCadImpTmp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                sCadImpTmp = sCadImpTmp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & Chr$(10)
                sCadImpTmp = sCadImpTmp & lsTab & String(130, "-") & Chr$(10)
                If nNumCab <= 55 Then
                    sCadImp = sCadImp & sCadImpTmp
                    sCadImpTmp = ""
                    nNumLinTmp = 0
                End If
            End If
        End If
        
        sCadImpTmp = sCadImpTmp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 30, 0) & _
                        ImpreFormat(Format(R!dRefinanciado, "dd/mm/yyyy"), 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & ImpreFormat(R!nInteresComp + R!nInteresMor + R!nGastos, 12) & Chr$(10)
        nNumCab = nNumCab + 1
        nNumLinTmp = nNumLinTmp + 1
        
        If nNumCab <= 55 Then
            sCadImp = sCadImp & sCadImpTmp
            sCadImpTmp = ""
            nNumLinTmp = 0
        End If
        
        If nNumCab > 55 Then
            sCadImp = sCadImp & Chr$(12)
            Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "INTERESES EN SUSPENSO DE CREDITOS REFINANCIADOS", psNomCmac, , True)
            sCadImp = sCadImp & Chr$(10)
            sCadImp = sCadImp & Chr$(10)
            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & Chr$(10)
            sCadImp = sCadImp & Chr$(10)
            sCadImp = sCadImp & Chr$(10)
            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & Chr$(10)
            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & Chr$(10)
            sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
            sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
            sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA          I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & Chr$(10)
            sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
            sCadImp = sCadImp & sCadImpTmp
            nNumCab = nNumLinTmp + 15
            sCadImpTmp = ""
        End If
        
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
Dim lnLineas As Long
Dim oNCred As NCredito

    Set oNCred = New NCredito
    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosRefinanciados(pMatAgencias, pdFecIni, pnMoneda, pMatProd)
    Set oDCred = Nothing
    sCadImp = ""
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS REFINANCIADOS", psNomCmac, , True)
        lnLineas = 5
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & Chr$(10)
        sCadImp = sCadImp & Chr$(10)
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & Chr$(10)
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & Chr$(10)
        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
        sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA  CAP INT   CAPITAL     I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & Chr$(10)
        sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
        lnLineas = lnLineas + 9
    End If
    Do While Not R.EOF
        'FINAL DE FUENTE DE FINANCIAMIENTO
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL FINANCIAMIENTO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & Chr$(10)
            sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
            sCadImp = sCadImp & Chr$(10)
            lnLineas = lnLineas + 4
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE PLAZO
            If Trim(R!sFuenteFinan) = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PLAZO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                   ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & Chr$(10)
                sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                sCadImp = sCadImp & Chr$(10)
                lnLineas = lnLineas + 4
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DE PRODUCTO
                If Trim(R!sPlazo) = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL PRODUCTO ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & Chr$(10)
                    sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                    sCadImp = sCadImp & Chr$(10)
                    lnLineas = lnLineas + 4
                    R.MoveNext
                    If R.EOF Then
                        Exit Do
                    End If
                    'FINAL DE AGENCIA
                    If R!sProducto = "TOTAL" Then
                        sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                        sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & Chr$(10)
                        sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                        sCadImp = sCadImp & Chr$(10)
                        lnLineas = lnLineas + 4
                        R.MoveNext
                        If R.EOF Then
                            Exit Do
                        End If
                        'FINAL
                        If R!sAgencia = "TOTAL" Then
                            sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 31) & ImpreFormat("CASOS : " & R!nCasos, 33, 0) & _
                                    ImpreFormat(R!nCapital, 12) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & Chr$(10)
                            sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                            sCadImp = sCadImp & Chr$(10)
                            lnLineas = lnLineas + 4
                            Exit Do
                        Else
                            nPuntPag = 0
                            sCadImp = sCadImp & Chr$(10)
                            sCadImp = sCadImp & Chr$(10)
                            sCadImp = sCadImp & lsTab & " Moneda : " & IIf(pnMoneda = gMonedaNacional, "SOLES", "DOLARES") & Chr$(10)
                            sCadImp = sCadImp & Chr$(10)
                            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & Chr$(10)
                            sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & Chr$(10)
                            sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
                            sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                            sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                            sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA  CAP INT   CAPITAL     I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & Chr$(10)
                            sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                            lnLineas = lnLineas + 11
                        End If
                    Else
                        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!sProducto) & Chr$(10)
                        sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
                        sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                        sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                        sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA  CAP INT   CAPITAL     I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & Chr$(10)
                        sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                        lnLineas = lnLineas + 6
                    End If
                Else
                    sCadImp = sCadImp & lsTab & "PLAZO : " & Trim(R!sPlazo) & Chr$(10)
                    sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                    sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                    sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA  CAP INT   CAPITAL     I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & Chr$(10)
                    sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                    lnLineas = lnLineas + 5
                End If
            Else
                sCadImp = sCadImp & lsTab & "FINANCIAMIENTO : " & Trim(R!sFuenteFinan) & Chr$(10)
                sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA  CAP INT   CAPITAL     I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & Chr$(10)
                sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
                lnLineas = lnLineas + 4
            End If
        End If
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 30, 0) & _
                        ImpreFormat(Format(R!dRefinanciado, "dd/mm/yyyy"), 12) & ImpreFormat(IIf(oNCred.EsRefinanciado(R!cCtaCod, True), "SI", "NO"), 2, 0) & ImpreFormat(R!nCapital, 10) & ImpreFormat(R!nInteresComp, 12) & ImpreFormat(R!nInteresMor, 12) & ImpreFormat(R!nGastos, 12) & Chr$(10)
        
        lnLineas = lnLineas + 1
        
        If lnLineas >= 60 Then
            sCadImp = sCadImp & Chr(12)
            Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS REFINANCIADOS", psNomCmac, , True)
            lnLineas = 4
            sCadImp = sCadImp & Chr$(10)
            sCadImp = sCadImp & Chr$(10)
            sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
            sCadImp = sCadImp & lsTab & ImpreFormat("  CREDITO              CLIENTE                           FECHA  CAP INT   CAPITAL     I. COMPENS.      I. MORAT.        GASTOS       DEUDA   ", 130) & Chr$(10)
            sCadImp = sCadImp & lsTab & String(130, "-") & Chr$(10)
            lnLineas = lnLineas + 5
        End If
        
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeCreditosRefinanciados = sCadImp
End Function

Public Function ImpLavDineroM(ByVal psFechaIni As String, ByVal psFechaFin As String, ByVal psNomCmac As String, _
                              ByVal psNomAge As String, ByVal psAgencias As String, pdFecSis As Date, ByVal psCodUser As String, ByVal psMonedas As String) As String


Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset
Dim nTipCam As Double
Dim nLineas As Integer
Dim sCodTemp As String
Dim nItem As Long
Dim lsImpre As String
Dim lsAgencia As String * 15, lsFechaTrans As String * 19
Dim lsProducto As String * 10, lsOperacion As String * 15
Dim lnMontoES As Double
Dim lnMontoESD As Double
Dim lnMontoED As Double
Dim nTotal1 As Double
Dim nTotal2 As Double
Dim nTotal3 As Double
Dim nTotal4 As Double
Dim nTTotal1 As Double
Dim nTTotal2 As Double
Dim nTTotal3 As Double
Dim nTTotal4 As Double


    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaImpLavDineroM(psFechaIni, psFechaFin, psAgencias, psMonedas)
    nTipCam = oDCred.RecuperaTCLavDineroM(psFechaIni)
    
    Set oDCred = Nothing
    sCadImp = ""

    If R.EOF Then
    Else
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "OPERACIONES DE CREDITOS CONTROLADAS CON LAVADO DE DINERO", psNomCmac, , True)
        nLineas = 5
        sCadImp = sCadImp & Space(80) & "TIPO DE CAMBIO :   " & nTipCam & Chr(10) & Chr(10)
        sCadImp = sCadImp & R!cPersCod & " " & R!cpersnombre & lnSaltoLinDoc
        sCadImp = sCadImp & String(180, "-") & Chr(10)
        sCadImp = sCadImp & "ITEM" & Space(3) & "AGENCIA" & Space(13) & "FECHA TRANSACCION" & Space(5) & "PRODUCTO"
        sCadImp = sCadImp & Space(8) & "OPERACION" & Space(15) & "CUENTA" & Space(14) & "MONTO SOLES" & Space(5)
        sCadImp = sCadImp & "EQUIV. DOLAR" & Space(5) & "MONTO DOLAR" & Space(5) & "TOTAL DOLAR" & Space(4) & "USUARIO" & Chr(10)
        sCadImp = sCadImp & String(180, "-") & Chr(10)
        nLineas = nLineas + 6
    End If
    Do While Not R.EOF
        If sCodTemp <> R!cPersCod And Len(Trim(sCodTemp)) > 0 Then
            sCadImp = sCadImp & String(180, "-") & Chr(10)
            sCadImp = sCadImp & "TOTAL CLIENTE" & Space(95) & ImpreFormat(nTotal1, 11, 2, True) & Space(3) & ImpreFormat(nTotal2, 11, 2, True) & Space(3) & ImpreFormat(nTotal3, 11, 2, True) & Space(3) & ImpreFormat(nTotal4, 11, 2, True) & Chr(10)
            sCadImp = sCadImp & String(180, "-") & Chr(10)
            nLineas = nLineas + 3
            nItem = 0
            nTotal1 = 0
            nTotal2 = 0
            nTotal3 = 0
            nTotal4 = 0
            sCodTemp = R!cPersCod
            
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & R!cPersCod & " " & R!cpersnombre & lnSaltoLinDoc
            sCadImp = sCadImp & String(180, "-") & Chr(10)
            sCadImp = sCadImp & "ITEM" & Space(3) & "AGENCIA" & Space(13) & "FECHA TRANSACCION" & Space(5) & "PRODUCTO"
            sCadImp = sCadImp & Space(8) & "OPERACION" & Space(15) & "CUENTA" & Space(14) & "MONTO SOLES" & Space(5)
            sCadImp = sCadImp & "EQUIV. DOLAR" & Space(5) & "MONTO DOLAR" & Space(5) & "TOTAL DOLAR" & Space(4) & "USUARIO" & Chr(10)
            sCadImp = sCadImp & String(180, "-") & Chr(10)
            nLineas = nLineas + 6
        End If
        If Len(Trim(sCodTemp)) = 0 Then
            sCodTemp = R!cPersCod
        End If
        
        nItem = nItem + 1
        lsAgencia = R!cAgeDescripcion
        lsFechaTrans = R!cFechaHora
        lsProducto = R!cDesProducto
        lsOperacion = R!cOpeDesc
        lnMontoES = 0
        lnMontoESD = 0
        lnMontoED = 0
        If R!cMoneda = "1" Then
            lnMontoES = R!nMonto
            lnMontoESD = R!nMonto / nTipCam
        Else
            lnMontoED = R!nMonto
        End If
        
        nTotal1 = nTotal1 + lnMontoES
        nTotal2 = nTotal2 + lnMontoESD
        nTotal3 = nTotal3 + lnMontoED
        nTotal4 = nTotal4 + lnMontoESD + lnMontoED
        
        nTTotal1 = nTTotal1 + lnMontoES
        nTTotal2 = nTTotal2 + lnMontoESD
        nTTotal3 = nTTotal3 + lnMontoED
        nTTotal4 = nTTotal4 + lnMontoESD + lnMontoED
        
        sCadImp = sCadImp & ImpreFormat(nItem, 3, 0) & Space(3) & lsAgencia & Space(5) & lsFechaTrans
        sCadImp = sCadImp & Space(5) & lsProducto & Space(5) & lsOperacion & Space(5) & R!cCtaCod & Space(5)
        sCadImp = sCadImp & ImpreFormat(lnMontoES, 11, 2, True) & Space(3) & ImpreFormat(lnMontoESD, 11, 2, True)
        sCadImp = sCadImp & Space(3) & ImpreFormat(lnMontoED, 11, 2, True) & Space(3)
        sCadImp = sCadImp & ImpreFormat(lnMontoESD + lnMontoED, 11, 2, True) & Space(3) & R!cUser & Chr(10)
         
        nLineas = nLineas + 1
        
        If nLineas > 52 Then
            sCadImp = sCadImp & Chr(12)
            Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "OPERACIONES DE CREDITOS CONTROLADAS CON LAVADO DE DINERO", psNomCmac, , True)
            nLineas = 5
            sCadImp = sCadImp & Space(80) & "TIPO DE CAMBIO :   " & nTipCam & Chr(10) & Chr(10)
            sCadImp = sCadImp & String(180, "-") & Chr(10)
            nLineas = nLineas + 3
        End If
        R.MoveNext
    Loop
    sCadImp = sCadImp & String(180, "-") & Chr(10)
    sCadImp = sCadImp & "TOTAL CLIENTE" & Space(95) & ImpreFormat(nTotal1, 11, 2, True) & Space(3) & ImpreFormat(nTotal2, 11, 2, True) & Space(3) & ImpreFormat(nTotal3, 11, 2, True) & Space(3) & ImpreFormat(nTotal4, 11, 2, True) & Chr(10)
    sCadImp = sCadImp & String(180, "-") & Chr(10) & Chr(10)
    sCadImp = sCadImp & String(180, "-") & Chr(10)
    sCadImp = sCadImp & "TOTAL        " & Space(95) & ImpreFormat(nTTotal1, 11, 2, True) & Space(3) & ImpreFormat(nTTotal2, 11, 2, True) & Space(3) & ImpreFormat(nTTotal3, 11, 2, True) & Space(3) & ImpreFormat(nTTotal4, 11, 2, True) & Chr(10)
    sCadImp = sCadImp & String(180, "-") & Chr(10)
 
    ImpLavDineroM = sCadImp
    
End Function

Public Function ImprimePersPendienteDevCredPers(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "", _
    Optional ByVal MatInstitucion As Variant = "") As String

Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset
Dim nTipCam As Double
Dim nLineas As Integer
Dim I As Long
Dim lnTotal As Long
Dim lnTotMoneda As Currency
Dim lnTotAgencia As Currency
Dim lnTotInst As Currency
Dim lsMoneda As String
Dim lsAgencias As String
Dim lsInstitucion As String
Dim oExcell As Excel.Application

Set oDCred = New DCredDoc
Set R = oDCred.RecuperaPersDevCredConv(pMatAgencias, pnMoneda, MatInstitucion)
Set oDCred = Nothing
sCadImp = ""
lnTotMoneda = 0
lnTotAgencia = 0
lnTotInst = 0

lsMoneda = ""
lsAgencias = ""
lsInstitucion = ""

If R.RecordCount > 0 Then
    lnTotal = R.RecordCount
    
    RaiseEvent IniciaBarra(lnTotal)
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "LISTADO PENDIENTES DE DEVOLUCION CREDITOS CON CONVENIO", psNomCmac, , True, psCodRepo)
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " Moneda : " & R!cMoneda & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("REGISTRO", 15) & ImpreFormat("CLIENTE", 35) & ImpreFormat("N°DOCUMENTO", 20) & ImpreFormat("MONTO", 10) & ImpreFormat("USER", 6) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    
    Do While Not R.EOF
        If lsInstitucion <> R!cInstitucion And lsInstitucion <> "" Then
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL INSTITUCION ", 69) & ImpreFormat(lnTotInst, 10) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            lnTotInst = 0
            sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("REGISTRO", 15) & ImpreFormat("CLIENTE", 35) & ImpreFormat("N°DOCUMENTO", 20) & ImpreFormat("MONTO", 10) & ImpreFormat("USER", 6) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            lsInstitucion = R!cInstitucion
        Else
            lsInstitucion = R!cInstitucion
            
        End If
        
        If lsAgencias <> R!sAgencia And lsAgencias <> "" Then
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 69) & ImpreFormat(lnTotAgencia, 10) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            lnTotAgencia = 0
            
            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("REGISTRO", 15) & ImpreFormat("CLIENTE", 35) & ImpreFormat("N°DOCUMENTO", 20) & ImpreFormat("MONTO", 10) & ImpreFormat("USER", 6) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            lsAgencias = R!sAgencia
        Else
            lsAgencias = R!sAgencia
            
        End If
        
        If lsMoneda <> R!cMoneda And lsMoneda <> "" Then
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL MONEDA", 69) & ImpreFormat(lnTotMoneda, 10) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            lnTotMoneda = 0
            
            sCadImp = sCadImp & lsTab & " Moneda : " & R!cMoneda & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("REGISTRO", 15) & ImpreFormat("CLIENTE", 35) & ImpreFormat("N°DOCUMENTO", 20) & ImpreFormat("MONTO", 10) & ImpreFormat("USER", 6) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            lsMoneda = R!cMoneda
        Else
            lsMoneda = R!cMoneda
            
        End If
        lnTotInst = lnTotInst + R!nMonto
        lnTotAgencia = lnTotAgencia + R!nMonto
        lnTotMoneda = lnTotMoneda + R!nMonto
        sCadImp = sCadImp & lsTab & ImpreFormat(Format(R!dRegistro, "dd/mm/yyyy"), 15, 0) & ImpreFormat(R!cTitular, 35) & ImpreFormat(R!cNroDoc, 15) & ImpreFormat(R!nMonto, 12, 2) & ImpreFormat(R!cUser, 6, 6) & lnSaltoLinDoc
        I = I + 1
        RaiseEvent ProgresoBarra(I, "Reporte de Pendientes por Devolver", R!cInstitucion)
        
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL INSTITUCION ", 69) & ImpreFormat(lnTotInst, 10) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 69) & ImpreFormat(lnTotAgencia, 10) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL MONEDA", 69) & ImpreFormat(lnTotMoneda, 10) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    RaiseEvent FinalizaBarra
    
    ImprimePersPendienteDevCredPers = sCadImp
Else
    ImprimePersPendienteDevCredPers = ""
End If

End Function

Public Function ImprimePersPendienteDevCredPersExcel(ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "", _
    Optional ByVal MatInstitucion As Variant = "") As String

Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset
Dim nTipCam As Double
Dim nLineas As Integer
Dim I As Long
Dim lnTotal As Long
Dim lnTotMoneda As Currency
Dim lnTotAgencia As Currency
Dim lnTotInst As Currency
Dim lsMoneda As String
Dim lsAgencias As String
Dim lsInstitucion As String
Dim vExcelObj As Excel.Application

Set oDCred = New DCredDoc
Set R = oDCred.RecuperaPersDevCredConv(pMatAgencias, pnMoneda, MatInstitucion)
Set oDCred = Nothing
sCadImp = ""
lnTotMoneda = 0
lnTotAgencia = 0
lnTotInst = 0

lsMoneda = ""
lsAgencias = ""
lsInstitucion = ""
Dim lnContIns As Long
Dim vIni As Long
Dim vItem As Long
Dim vCel As String
lnContIns = 0
If R.RecordCount > 0 Then
    lnTotal = R.RecordCount
    
    RaiseEvent IniciaBarra(lnTotal)
    
    Dim vNHC As String

    
    vNHC = App.path & "\spooler\Rep_Pend_Dev" & Format(pdFecSis, "yyyymmdd") & ".XLS"
    
    Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
    vExcelObj.DisplayAlerts = False
    
    vExcelObj.Workbooks.Add
    vExcelObj.Sheets("Hoja1").Select
    vExcelObj.Sheets("Hoja1").Name = "Rep_Dev_Cred"
    lnContIns = lnContIns + 1
    
    vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
    vExcelObj.Range("A1:IV65536").Font.Size = 8
    
    vExcelObj.Range("A1").Select
    vExcelObj.Range("A1").Font.Bold = True
    vExcelObj.Range("A1").HorizontalAlignment = 1
    vExcelObj.ActiveCell.value = psNomAge
    
    vExcelObj.Range("C3").Select
    vExcelObj.Range("C3").Font.Bold = True
    vExcelObj.Range("C3").HorizontalAlignment = 1
    vExcelObj.ActiveCell.value = "LISTADO PENDIENTES DE DEVOLUCION CREDITOS CON CONVENIO"
    
    vExcelObj.Range("G1").Select
    vExcelObj.Range("G1").Font.Bold = True
    vExcelObj.Range("G1").HorizontalAlignment = 1
    vExcelObj.ActiveCell.value = "Informacion al : " & Format(pdFecSis, "dd/mm/yyyy")
    
    vExcelObj.Range("A6").Select
    vExcelObj.Range("A6").Font.Bold = True
    vExcelObj.Range("A6").ColumnWidth = 40
    vExcelObj.ActiveCell.value = "INSTITUCION"
    
    vExcelObj.Range("B6").Select
    vExcelObj.Range("B6").Font.Bold = True
    vExcelObj.Range("B6").ColumnWidth = 10
    vExcelObj.ActiveCell.value = "REGISTRO"
    
    vExcelObj.Range("C6").Select
    vExcelObj.Range("C6").Font.Bold = True
    vExcelObj.Range("C6").ColumnWidth = 30
    vExcelObj.ActiveCell.value = "NOMBRE DE CLIENTE"
    
    vExcelObj.Range("D6").Select
    vExcelObj.Range("D6").Font.Bold = True
    vExcelObj.Range("D6").ColumnWidth = 10
    vExcelObj.ActiveCell.value = "DOCUMENTO"
    
    vExcelObj.Range("E6").Select
    vExcelObj.Range("E6").Font.Bold = True
    vExcelObj.Range("E6").ColumnWidth = 10
    vExcelObj.ActiveCell.value = "IMPORTE"
    
    vExcelObj.Range("F6").Select
    vExcelObj.Range("F6").Font.Bold = True
    vExcelObj.Range("F6").ColumnWidth = 10
    vExcelObj.ActiveCell.value = "USER"

    vExcelObj.Range("G6").Select
    vExcelObj.Range("G6").Font.Bold = True
    vExcelObj.Range("G6").ColumnWidth = 10
    vExcelObj.ActiveCell.value = "AGENCIA"
    
    vExcelObj.Range("H6").Select
    vExcelObj.Range("H6").Font.Bold = True
    vExcelObj.Range("H6").ColumnWidth = 10
    vExcelObj.ActiveCell.value = "MONEDA"
    
    vIni = 6
    vItem = vIni
    Do While Not R.EOF
        vItem = vItem + 1
            
        vCel = "A" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + Trim(R!cInstitucion)
        
        vCel = "B" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + Format(R!dRegistro, "dd/mm/yyyy")
            
        vCel = "C" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + Trim(R!cTitular)
        
        vCel = "D" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + Trim(R!cNroDoc)
        
        vCel = "E" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = ImpreFormat(R!nMonto, 12, 2)
        
        vCel = "F" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + R!cUser
        
        vCel = "G" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + R!sAgencia
        
        vCel = "H" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + R!cMoneda
        
        I = I + 1
        RaiseEvent ProgresoBarra(I, "Reporte de Pendientes por Devolver", R!cInstitucion)
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
        
    vItem = vItem + 1
    vExcelObj.Range("A6").Select
    'vExcelObj.Range("A6").Activate
    vExcelObj.Range("A6").SubTotal 1, xlSum, Array(5)
    'vExcelObj.Range("A" & vItem).Select
    'vExcelObj.Range("A" & vItem).Font.Bold = True
    'vExcelObj.Range("A" & vItem).ColumnWidth = 10
    'vExcelObj.ActiveCell.value = "TOTAL"
    
    'vExcelObj.Range("D" & vItem).Select
    'vExcelObj.Range("D" & vItem).Font.Bold = True
    'vExcelObj.Range("D" & vItem).ColumnWidth = 10
    'vExcelObj.ActiveCell.Formula = "=SUM(D" & vIni + 1 & ":D" & vItem - 1 & ")"
    
    RaiseEvent FinalizaBarra

    vExcelObj.Range("A1").Select
    vExcelObj.ActiveWorkbook.SaveAs (vNHC)
    vExcelObj.ActiveWorkbook.Close
    vExcelObj.Workbooks.Open (vNHC)
    vExcelObj.Visible = True
    
    Set vExcelObj = Nothing
    
    ImprimePersPendienteDevCredPersExcel = ""
Else
    ImprimePersPendienteDevCredPersExcel = ""
End If

End Function

Public Function ImprimeDevolucionCredPers(ByVal pdFecIni As Date, pdFecFin As Date, ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "", _
    Optional ByVal MatInstitucion As Variant = "", Optional pCastDev As Integer = 0) As String
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset
Dim nTipCam As Double
Dim nLineas As Integer
Dim I As Long
Dim lnTotal As Long
Dim lnTotMoneda As Currency
Dim lnTotAgencia As Currency
Dim lnTotInst As Currency
Dim lsMoneda As String
Dim lsAgencias As String
Dim lsInstitucion As String

Set oDCred = New DCredDoc
Set R = oDCred.RecuperaDevCredConv(pdFecIni, pdFecFin, pMatAgencias, pnMoneda, MatInstitucion, pCastDev)
Set oDCred = Nothing
sCadImp = ""
lnTotMoneda = 0
lnTotAgencia = 0
lnTotInst = 0

lsMoneda = ""
lsAgencias = ""
lsInstitucion = ""

If R.RecordCount > 0 Then
    lnTotal = R.RecordCount
    
    RaiseEvent IniciaBarra(lnTotal)
    
    If pCastDev = 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "LISTADO DE DEVOLUCION POR CREDITOS CON CONVENIO", psNomCmac, , True, psCodRepo)
    Else
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "LISTADO DE CASTIGO DE DEVOLUCION POR CREDITOS CON CONVENIO", psNomCmac, , True, psCodRepo)
    End If
    sCadImp = sCadImp & CentrarCadena("DESDE :" & Format(pdFecIni, "dd/mm/yyyy") & " HASTA " & Format(pdFecFin, "dd/mm/yyyy"), 120, 8)
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & " Moneda : " & R!cMoneda & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!cAgencia) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
     If pCastDev = 1 Then
        sCadImp = sCadImp & lsTab & ImpreFormat("FEC CASTIGO", 15) & ImpreFormat("REGISTRO", 15) & ImpreFormat("CLIENTE", 35) & ImpreFormat("N°DOCUMENTO", 23) & ImpreFormat("MONTO", 8) & ImpreFormat("USER", 6, 0) & lnSaltoLinDoc
     Else
        sCadImp = sCadImp & lsTab & ImpreFormat("DEVOLUCION", 15) & ImpreFormat("REGISTRO", 15) & ImpreFormat("CLIENTE", 35) & ImpreFormat("N°DOCUMENTO", 23) & ImpreFormat("MONTO", 8) & ImpreFormat("USER", 6, 0) & lnSaltoLinDoc
     End If
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    
    Do While Not R.EOF
        If lsInstitucion <> R!cInstitucion And lsInstitucion <> "" Then
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL INSTITUCION ", 88) & ImpreFormat(lnTotInst, 10) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            lnTotInst = 0
            sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            If pCastDev = 1 Then
                sCadImp = sCadImp & lsTab & ImpreFormat("FEC CASTIGO", 15) & ImpreFormat("REGISTRO", 15) & ImpreFormat("CLIENTE", 35) & ImpreFormat("N°DOCUMENTO", 23) & ImpreFormat("MONTO", 8) & ImpreFormat("USER", 6, 0) & lnSaltoLinDoc
            Else
                sCadImp = sCadImp & lsTab & ImpreFormat("DEVOLUCION", 15) & ImpreFormat("REGISTRO", 15) & ImpreFormat("CLIENTE", 35) & ImpreFormat("N°DOCUMENTO", 23) & ImpreFormat("MONTO", 8) & ImpreFormat("USER", 6, 0) & lnSaltoLinDoc
            End If
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            lsInstitucion = R!cInstitucion
        Else
            lsInstitucion = R!cInstitucion
        End If
        
        If lsAgencias <> R!cAgencia And lsAgencias <> "" Then
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 88) & ImpreFormat(lnTotAgencia, 10) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            lnTotAgencia = 0
            
            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!cAgencia) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            If pCastDev = 1 Then
                sCadImp = sCadImp & lsTab & ImpreFormat("FEC CASTIGO", 15) & ImpreFormat("REGISTRO", 15) & ImpreFormat("CLIENTE", 35) & ImpreFormat("N°DOCUMENTO", 23) & ImpreFormat("MONTO", 8) & ImpreFormat("USER", 6, 0) & lnSaltoLinDoc
            Else
                sCadImp = sCadImp & lsTab & ImpreFormat("DEVOLUCION", 15) & ImpreFormat("REGISTRO", 15) & ImpreFormat("CLIENTE", 35) & ImpreFormat("N°DOCUMENTO", 23) & ImpreFormat("MONTO", 8) & ImpreFormat("USER", 6, 0) & lnSaltoLinDoc
            End If
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            lsAgencias = R!cAgencia
        Else
            lsAgencias = R!cAgencia
        End If
        
        If lsMoneda <> R!cMoneda And lsMoneda <> "" Then
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL MONEDA", 88) & ImpreFormat(lnTotMoneda, 10) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            lnTotMoneda = 0
            
            sCadImp = sCadImp & lsTab & " Moneda : " & R!cMoneda & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            If pCastDev = 1 Then
                sCadImp = sCadImp & lsTab & ImpreFormat("FEC CASTIGO", 15) & ImpreFormat("REGISTRO", 15) & ImpreFormat("CLIENTE", 35) & ImpreFormat("N°DOCUMENTO", 23) & ImpreFormat("MONTO", 8) & ImpreFormat("USER", 6, 0) & lnSaltoLinDoc
            Else
                sCadImp = sCadImp & lsTab & ImpreFormat("DEVOLUCION", 15) & ImpreFormat("REGISTRO", 15) & ImpreFormat("CLIENTE", 35) & ImpreFormat("N°DOCUMENTO", 23) & ImpreFormat("MONTO", 8) & ImpreFormat("USER", 6, 0) & lnSaltoLinDoc
            End If
            sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
            lsMoneda = R!cMoneda
        Else
            lsMoneda = R!cMoneda
            
        End If
        lnTotAgencia = lnTotAgencia + R!nMonto
        lnTotMoneda = lnTotMoneda + R!nMonto
        lnTotInst = lnTotInst + R!nMonto
        sCadImp = sCadImp & lsTab & ImpreFormat(R!DOperacion, 15) & ImpreFormat(R!dRegistro, 15) & ImpreFormat(R!cTitular, 35) & ImpreFormat(R!cNroDoc, 15) & ImpreFormat(R!nMonto, 12, 2, True) & ImpreFormat(R!cUser, 6, 4) & lnSaltoLinDoc
        I = I + 1
        RaiseEvent ProgresoBarra(I, "Reporte de Pendientes por Devolver", R!cInstitucion)
        
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL INSTITUCION ", 88) & ImpreFormat(lnTotInst, 10) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 88) & ImpreFormat(lnTotAgencia, 10) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL MONEDA", 88) & ImpreFormat(lnTotMoneda, 10) & lnSaltoLinDoc
    sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    RaiseEvent FinalizaBarra
    
    ImprimeDevolucionCredPers = sCadImp
Else
    ImprimeDevolucionCredPers = ""
End If

End Function
Public Function ImprimeDevolucionCredPersExcel(ByVal pdFecIni As Date, pdFecFin As Date, ByVal pMatAgencias As Variant, _
    ByVal pnMoneda As Moneda, ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "", _
    Optional ByVal MatInstitucion As Variant = "", Optional ByVal pnCastigoDev As Integer = 0) As String
Dim oDCred As DCredDoc
Dim sCadImp As String
Dim R As ADODB.Recordset
Dim nTipCam As Double
Dim nLineas As Integer
Dim I As Long
Dim lnTotal As Long
Dim lnTotMoneda As Currency
Dim lnTotAgencia As Currency
Dim lnTotInst As Currency
Dim lsMoneda As String
Dim lsAgencias As String
Dim lsInstitucion As String
Dim vExcelObj As Excel.Application
Dim lnContIns As Long
Dim vIni As Long
Dim vItem As Long
Dim vCel As String
Dim vNHC As String

Set oDCred = New DCredDoc
'MADM 20110930 pnPagoCastigo
Set R = oDCred.RecuperaDevCredConv(pdFecIni, pdFecFin, pMatAgencias, pnMoneda, MatInstitucion, pnCastigoDev)
Set oDCred = Nothing
sCadImp = ""
lnTotMoneda = 0
lnTotAgencia = 0
lnTotInst = 0

lsMoneda = ""
lsAgencias = ""
lsInstitucion = ""

If R.RecordCount > 0 Then
    lnTotal = R.RecordCount
    
    RaiseEvent IniciaBarra(lnTotal)
    
    vNHC = App.path & "\spooler\Rep_Dev_CredPers.XLS"
    
    Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
    vExcelObj.DisplayAlerts = False
    
    vExcelObj.Workbooks.Add
    vExcelObj.Sheets("Hoja1").Select
    vExcelObj.Sheets("Hoja1").Name = "Rep_Dev_Cred"
    lnContIns = lnContIns + 1
    
    vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
    vExcelObj.Range("A1:IV65536").Font.Size = 8
    
    vExcelObj.Range("A1").Select
    vExcelObj.Range("A1").Font.Bold = True
    vExcelObj.Range("A1").HorizontalAlignment = 1
    vExcelObj.ActiveCell.value = psNomAge
    
    vExcelObj.Range("C3").Select
    vExcelObj.Range("C3").Font.Bold = True
    vExcelObj.Range("C3").HorizontalAlignment = 1
    'madm 20110930
    If pnCastigoDev = 1 Then
        vExcelObj.ActiveCell.value = "LISTADO DE CASTIGO DEVOLUCION POR CREDITOS CON CONVENIO "
    Else
        vExcelObj.ActiveCell.value = "LISTADO DE DEVOLUCION POR CREDITOS CON CONVENIO "
    End If
    'END MADM
    vExcelObj.Range("C4").Select
    vExcelObj.Range("C4").Font.Bold = True
    vExcelObj.Range("C4").HorizontalAlignment = 1
    vExcelObj.ActiveCell.value = "DESDE :" & Format(pdFecIni, "dd/mm/yyyy") & " HASTA " & Format(pdFecFin, "dd/mm/yyyy")
    
    vExcelObj.Range("A6").Select
    vExcelObj.Range("A6").Font.Bold = True
    vExcelObj.Range("A6").ColumnWidth = 40
    vExcelObj.ActiveCell.value = "INSTITUCION"
    
    vExcelObj.Range("B6").Select
    vExcelObj.Range("B6").Font.Bold = True
    vExcelObj.Range("B6").ColumnWidth = 10
    If pnCastigoDev = 1 Then
        vExcelObj.ActiveCell.value = "CASTIGO"
    Else
        vExcelObj.ActiveCell.value = "DEVOLUCION"
    End If
    
    vExcelObj.Range("C6").Select
    vExcelObj.Range("C6").Font.Bold = True
    vExcelObj.Range("C6").ColumnWidth = 10
    vExcelObj.ActiveCell.value = "REGISTRO"
    
    vExcelObj.Range("D6").Select
    vExcelObj.Range("D6").Font.Bold = True
    vExcelObj.Range("D6").ColumnWidth = 30
    vExcelObj.ActiveCell.value = "NOMBRE DE CLIENTE"
    
    vExcelObj.Range("E6").Select
    vExcelObj.Range("E6").Font.Bold = True
    vExcelObj.Range("E6").ColumnWidth = 10
    vExcelObj.ActiveCell.value = "DOCUMENTO"
    
    vExcelObj.Range("F6").Select
    vExcelObj.Range("F6").Font.Bold = True
    vExcelObj.Range("F6").ColumnWidth = 10
    vExcelObj.ActiveCell.value = "IMPORTE"
    
    vExcelObj.Range("G6").Select
    vExcelObj.Range("G6").Font.Bold = True
    vExcelObj.Range("G6").ColumnWidth = 10
    vExcelObj.ActiveCell.value = "USER"

    vExcelObj.Range("H6").Select
    vExcelObj.Range("H6").Font.Bold = True
    vExcelObj.Range("H6").ColumnWidth = 20
    vExcelObj.ActiveCell.value = "AGENCIA"
    
    vExcelObj.Range("I6").Select
    vExcelObj.Range("I6").Font.Bold = True
    vExcelObj.Range("I6").ColumnWidth = 10
    vExcelObj.ActiveCell.value = "MONEDA"
    
    vIni = 6
    vItem = vIni
    Do While Not R.EOF
        vItem = vItem + 1
            
        vCel = "A" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + Trim(R!cInstitucion)
        
        vCel = "B" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + Format(R!DOperacion, "dd/mm/yyyy")
        
        vCel = "C" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + Format(R!dRegistro, "dd/mm/yyyy")
            
        vCel = "D" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + Trim(R!cTitular)
        
        vCel = "E" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + Trim(R!cNroDoc)
        
        vCel = "F" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = ImpreFormat(R!nMonto, 12, 2)
        
        vCel = "G" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + R!cUser
        
        vCel = "H" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + R!cAgencia
        
        vCel = "I" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = "'" + R!cMoneda
        
        I = I + 1
        RaiseEvent ProgresoBarra(I, "Reporte de Pendientes por Devolver", R!cInstitucion)
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
        
    vItem = vItem + 1
    vExcelObj.Range("A6").Select
    vExcelObj.Range("A6").SubTotal 1, xlSum, Array(6)
    
    RaiseEvent FinalizaBarra

    vExcelObj.Range("A1").Select
    vExcelObj.ActiveWorkbook.SaveAs (vNHC)
    vExcelObj.ActiveWorkbook.Close
    vExcelObj.Workbooks.Open (vNHC)
    vExcelObj.Visible = True
    
    Set vExcelObj = Nothing
    
    
    ImprimeDevolucionCredPersExcel = sCadImp
Else
    ImprimeDevolucionCredPersExcel = ""
End If

End Function


Public Function ImprimeProtestoDia(ByVal psCodAgen As String, ByVal pdFechaFiltro As Date, _
ByVal psCodUser As String, ByVal pnNomCMAC As String, ByVal pdFecSis As Date) As String
    Dim oDCred As DCredDoc
    Dim sCadImp As String
    Dim rs As ADODB.Recordset
    Dim lnTotal As String
    Dim sNomAge As String
    Dim oVisualizacion As DVisualizacion
    Dim I As Integer
    Dim lsCodAge As String
    
    Set oDCred = New DCredDoc
    Set rs = oDCred.ListaProtestosDia(psCodAgen, Format(pdFechaFiltro, "dd/MM/yyyy"))
    Set oDCred = Nothing
    
    If rs.RecordCount > 0 Then
        lnTotal = rs.RecordCount
        
        RaiseEvent IniciaBarra(lnTotal)
        
        Set oVisualizacion = New DVisualizacion
        sNomAge = oVisualizacion.ObtenerNomAge(rs!CCodage)
        Set oVisualizacion = Nothing
        
        Call ImprimeCabeceraDocumento(sCadImp, sNomAge, pdFecSis, psCodUser, "LISTADO DE PROTESTOS : " & Format(pdFechaFiltro, "dd/mm/yyyy"), "CMAC ICA")
        sCadImp = sCadImp & Chr(10)
        sCadImp = sCadImp & Chr(10)
        sCadImp = sCadImp & Chr(10)
      
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO NUEVO", 20) & ImpreFormat("CREDITO ANTIGUO", 20) & ImpreFormat("CLIENTE ", 35) & _
            ImpreFormat("MONTO", 15) & ImpreFormat("MONEDA", 10) & lnSaltoLinDoc

        sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
        
        lsCodAge = rs!CCodage
        If Not rs.EOF And Not rs.BOF Then
            rs.MoveFirst
        End If
        Do Until rs.EOF
            If lsCodAge <> rs!CCodage Then
                sCadImp = sCadImp & Chr(12)
                Set oVisualizacion = New DVisualizacion
                sNomAge = oVisualizacion.ObtenerNomAge(rs!CCodage)
                Set oVisualizacion = Nothing
                
                Call ImprimeCabeceraDocumento(sCadImp, sNomAge, pdFecSis, psCodUser, "LISTADO DE PROTESTOS : " & Format(pdFechaFiltro, "dd/mm/yyyy"), "CMAC ICA")
                sCadImp = sCadImp & Chr(10)
                sCadImp = sCadImp & Chr(10)
                sCadImp = sCadImp & Chr(10)
              
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO NUEVO", 20) & ImpreFormat("CREDITO ANTIGUO", 20) & ImpreFormat("CLIENTE ", 35) & _
                ImpreFormat("MONTO", 15) & ImpreFormat("MONEDA", 10) & lnSaltoLinDoc
        
                sCadImp = sCadImp & lsTab & String(155, "-") & lnSaltoLinDoc
                lsCodAge = rs!CCodage
            End If
            
            sCadImp = sCadImp & lsTab & ImpreFormat(rs!cCtaCod, 20) & ImpreFormat(IIf(IsNull(rs!CctacodAnt), "", rs!CctacodAnt), 20) & ImpreFormat(rs!cpersnombre, 30) & _
                                  ImpreFormat(rs!nMonto, 10) & Space(10) & ImpreFormat(rs!Moneda, 10) & Chr(10)
            
            RaiseEvent ProgresoBarra(I, "Lista de creditos protestados ", "Del dia : " & pdFechaFiltro)
            I = I + 1
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    RaiseEvent FinalizaBarra
    
    ImprimeProtestoDia = sCadImp
End Function
Public Sub ConsolidadoPorProductoAna(ByVal psCodAgencia As String, ByVal pdFechaFiltro As Date, _
                                     ByVal pMatAgencias As Variant, ByVal psCodUser As String, ByVal pdFecSis As Date, _
                                     ByVal pnTipoCambio As Double, ByVal pbFecha As Boolean)

Dim oCredDoc As DCredDoc
Dim oVisualizacion As DVisualizacion
Dim sNomAge As String
Dim sAgencia As String
Dim sProducto As String
Dim rs As ADODB.Recordset
Dim rsVigente As ADODB.Recordset
Dim I As Integer
Dim oConec As DConecta
Dim nCodigo As Integer
Dim sSQL As String

    For I = 0 To UBound(pMatAgencias) - 1
        If I = 0 Then
            sAgencia = "'" & pMatAgencias(0) & "'"
        Else
             sAgencia = sAgencia & ",'" & pMatAgencias(I) & "'"
        End If
    Next I
    
    Set oVisualizacion = New DVisualizacion
    sNomAge = oVisualizacion.ObtenerNomAge(psCodAgencia)
    Set oVisualizacion = Nothing
    
    '--*******************************************
    Set oConec = New DConecta
    oConec.AbreConexion
    sSQL = "Select IsNull(Max(IdCodigo),0) as IdCodigo From TmpConsolidadoProducto"
    
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    
    If Not rs.EOF And Not rs.BOF Then
        nCodigo = rs!IdCodigo + 1
    End If
    Set rs = Nothing
    '--********************************************
    
    'Armando la data del reporte de la tabla tmpConsolidadoProdicto
    'Producto Pymes 201
   Set oCredDoc = New DCredDoc
   Set rs = oCredDoc.ListaConsolidadoPorAnalista(pnTipoCambio, sAgencia, pdFechaFiltro, psCodUser, "201", pbFecha, "201")
   Set oCredDoc = Nothing
   
   If rs.RecordCount > 0 Then
       Set oCredDoc = New DCredDoc
       Call oCredDoc.InsertarTmpConsolidadoProducto(rs, psCodUser, pdFechaFiltro, sProducto, sAgencia, pbFecha, "201", nCodigo)
       Set oCredDoc = Nothing
   End If
   
   'Producto Consumo=304
   Set oCredDoc = New DCredDoc
   Set rs = oCredDoc.ListaConsolidadoPorAnalista(pnTipoCambio, sAgencia, pdFechaFiltro, psCodUser, "304", pbFecha, "304")
   Set oCredDoc = Nothing
   
   If rs.RecordCount > 0 Then
       Set oCredDoc = New DCredDoc
       Call oCredDoc.InsertarTmpConsolidadoProducto(rs, psCodUser, pdFechaFiltro, sProducto, sAgencia, pbFecha, "304", nCodigo)
       Set oCredDoc = Nothing
   End If
   Set rs = Nothing
   
  'Producto DsctoxPlanilla=301
  Set oCredDoc = New DCredDoc
   Set rs = oCredDoc.ListaConsolidadoPorAnalista(pnTipoCambio, sAgencia, pdFechaFiltro, psCodUser, "301", pbFecha, "301")
   Set oCredDoc = Nothing
   
   If rs.RecordCount > 0 Then
       Set oCredDoc = New DCredDoc
       Call oCredDoc.InsertarTmpConsolidadoProducto(rs, psCodUser, pdFechaFiltro, sProducto, sAgencia, pbFecha, "301", nCodigo)
       Set oCredDoc = Nothing
   End If
   Set rs = Nothing
   
   'Producto CTS =303
  Set oCredDoc = New DCredDoc
   Set rs = oCredDoc.ListaConsolidadoPorAnalista(pnTipoCambio, sAgencia, pdFechaFiltro, psCodUser, "303", pbFecha, "303")
   Set oCredDoc = Nothing
   
   If rs.RecordCount > 0 Then
       Set oCredDoc = New DCredDoc
       Call oCredDoc.InsertarTmpConsolidadoProducto(rs, psCodUser, pdFechaFiltro, sProducto, sAgencia, pbFecha, "303", nCodigo)
       Set oCredDoc = Nothing
   End If
   Set rs = Nothing
   
  'Producto Plazo Fijo =302
   Set oCredDoc = New DCredDoc
   Set rs = oCredDoc.ListaConsolidadoPorAnalista(pnTipoCambio, sAgencia, pdFechaFiltro, psCodUser, "302", pbFecha, "302")
   Set oCredDoc = Nothing
   
   If rs.RecordCount > 0 Then
       Set oCredDoc = New DCredDoc
       Call oCredDoc.InsertarTmpConsolidadoProducto(rs, psCodUser, pdFechaFiltro, sProducto, sAgencia, pbFecha, "302", nCodigo)
       Set oCredDoc = Nothing
   End If
   Set rs = Nothing
   
   'Producto Funcionarios CMAC 320
   Set oCredDoc = New DCredDoc
   Set rs = oCredDoc.ListaConsolidadoPorAnalista(pnTipoCambio, sAgencia, pdFechaFiltro, psCodUser, "320", pbFecha, "320")
   Set oCredDoc = Nothing
   
   If rs.RecordCount > 0 Then
       Set oCredDoc = New DCredDoc
       Call oCredDoc.InsertarTmpConsolidadoProducto(rs, psCodUser, pdFechaFiltro, sProducto, sAgencia, pbFecha, "320", nCodigo)
       Set oCredDoc = Nothing
   End If
   Set rs = Nothing
   
   'Producto Mes Agricola 202
   Set oCredDoc = New DCredDoc
   Set rs = oCredDoc.ListaConsolidadoPorAnalista(pnTipoCambio, sAgencia, pdFechaFiltro, psCodUser, "202", pbFecha, "202")
   Set oCredDoc = Nothing
   
   If rs.RecordCount > 0 Then
       Set oCredDoc = New DCredDoc
       Call oCredDoc.InsertarTmpConsolidadoProducto(rs, psCodUser, pdFechaFiltro, sProducto, sAgencia, pbFecha, "202", nCodigo)
       Set oCredDoc = Nothing
   End If
   Set rs = Nothing
   
   'Producto Comercial Agricola Comercial Agricola 102
   Set oCredDoc = New DCredDoc
   Set rs = oCredDoc.ListaConsolidadoPorAnalista(pnTipoCambio, sAgencia, pdFechaFiltro, psCodUser, "102", pbFecha, "102")
   Set oCredDoc = Nothing
   
   If rs.RecordCount > 0 Then
       Set oCredDoc = New DCredDoc
       Call oCredDoc.InsertarTmpConsolidadoProducto(rs, psCodUser, pdFechaFiltro, sProducto, sAgencia, pbFecha, "102", nCodigo)
       Set oCredDoc = Nothing
   End If
   Set rs = Nothing
   
   'Producto Comercial Empresarial Comercial Agricola 101
   Set oCredDoc = New DCredDoc
   Set rs = oCredDoc.ListaConsolidadoPorAnalista(pnTipoCambio, sAgencia, pdFechaFiltro, psCodUser, "101", pbFecha, "101")
   Set oCredDoc = Nothing
   
   If rs.RecordCount > 0 Then
       Set oCredDoc = New DCredDoc
       Call oCredDoc.InsertarTmpConsolidadoProducto(rs, psCodUser, pdFechaFiltro, sProducto, sAgencia, pbFecha, "101", nCodigo)
       Set oCredDoc = Nothing
   End If
   Set rs = Nothing
   
   'Producto Hipotecario 423
   Set oCredDoc = New DCredDoc
   Set rs = oCredDoc.ListaConsolidadoPorAnalista(pnTipoCambio, sAgencia, pdFechaFiltro, psCodUser, "423", pbFecha, "423")
   Set oCredDoc = Nothing
   
   If rs.RecordCount > 0 Then
       Set oCredDoc = New DCredDoc
       Call oCredDoc.InsertarTmpConsolidadoProducto(rs, psCodUser, pdFechaFiltro, sProducto, sAgencia, pbFecha, "423", nCodigo)
       Set oCredDoc = Nothing
   End If
   Set rs = Nothing
   
   'Producto Hipotecario 305
   Set oCredDoc = New DCredDoc
   Set rs = oCredDoc.ListaConsolidadoPorAnalista(pnTipoCambio, sAgencia, pdFechaFiltro, psCodUser, "305", pbFecha, "305")
   Set oCredDoc = Nothing
   
   If rs.RecordCount > 0 Then
       Set oCredDoc = New DCredDoc
       Call oCredDoc.InsertarTmpConsolidadoProducto(rs, psCodUser, pdFechaFiltro, sProducto, sAgencia, pbFecha, "305", nCodigo)
       Set oCredDoc = Nothing
   End If
   Set rs = Nothing
   
End Sub

Public Function ConsolidadoPorAnalista(ByVal psCodAgencia As String, ByVal pdFecha As Date, pnTipoCambio As Double, _
                                       ByVal pMatAgencias As Variant, ByVal psCodUser As String, ByVal pMatAnalistas As Variant, _
                                       ByVal pMatProducto As Variant, ByVal pbFecha As Boolean) As String
    Dim odCredDoc As DCredDoc
    Dim oVisualizacion As DVisualizacion
    
    Dim sCadImp As String
    Dim sNomAge As String
    Dim sAgencia As String
    Dim sAnalistas As String
    Dim sProducto As String
    
    Dim rs As ADODB.Recordset
    Dim rsVigente As ADODB.Recordset
    
    Dim nTotal As Integer
    Dim I As Integer
    
    Dim sUser As String
    
    sCadImp = ""
    
    For I = 0 To UBound(pMatProducto) - 1
        If I = 0 Then
            sProducto = "'" & pMatProducto(I) & "'"
        Else
            sProducto = sProducto & ",'" & pMatProducto(I) & "'"
        End If
    Next I
    
    
    For I = 0 To UBound(pMatAgencias) - 1
        If I = 0 Then
            sAgencia = "'" & pMatAgencias(0) & "'"
        Else
            sAgencia = sAgencia & ",'" & pMatAgencias(I) & "'"
        End If
    Next I
    
    For I = 0 To UBound(pMatAnalistas) - 1
        If I = 0 Then
            Set odCredDoc = New DCredDoc
            sUser = "'" & odCredDoc.ListaUser(pMatAnalistas(0)) & "'"
            Set odCredDoc = Nothing
            
            sAnalistas = "'" & pMatAnalistas(0) & "'"
        Else
            Set odCredDoc = New DCredDoc
            sUser = sUser & ",'" & odCredDoc.ListaUser(pMatAnalistas(I)) & "'"
            Set odCredDoc = Nothing
            sAnalistas = sAnalistas & ",'" & pMatAnalistas(I) & "'"
        End If
    Next I
    
    Set oVisualizacion = New DVisualizacion
    sNomAge = oVisualizacion.ObtenerNomAge(psCodAgencia)
    Set oVisualizacion = Nothing
    
    Set odCredDoc = New DCredDoc
    Set rs = odCredDoc.ListaConsolidadoPorAnalista(pnTipoCambio, sAgencia, pdFecha, sUser, sProducto, pbFecha)
    Set odCredDoc = Nothing
    
    If rs.RecordCount > 0 Then
        nTotal = rs.RecordCount
        RaiseEvent IniciaBarra(nTotal)
        Set odCredDoc = New DCredDoc
        Call odCredDoc.InsertarTmpConsolidadoAnalistas(rs, psCodUser, pdFecha, sProducto, sAgencia, pbFecha)
        Set odCredDoc = Nothing
        
        For I = 1 To nTotal
            RaiseEvent ProgresoBarra(I, "Consolidado por analista", "Procesando")
        Next I
     
        RaiseEvent FinalizaBarra
    Else
        Exit Function
    End If
    
End Function
Public Function ConsolidadoAnalistaV2(psCodAgencia As String, psCodUsuario As String, _
ByVal pdFecha As Date, ByVal pnTipoCambio As Double, ByVal pbFecha As Boolean) As String
    Dim odCredDoc As DCredDoc
    Dim rs As ADODB.Recordset
    Dim rsVigente As ADODB.Recordset
    
    Dim nTotal As Integer
    Dim I As Integer
    Dim sCadImp As String
    
    sCadImp = ""
    
    Set odCredDoc = New DCredDoc
    Set rs = odCredDoc.ListaConsolidadoPorAnalistaV2(pnTipoCambio, psCodAgencia, pdFecha, psCodUsuario, pbFecha)
    Set odCredDoc = Nothing
    
    If rs.RecordCount > 0 Then
        nTotal = rs.RecordCount
        RaiseEvent IniciaBarra(nTotal)
        Set odCredDoc = New DCredDoc
        Call odCredDoc.InsertarTmpConsolidadoAnalistas(rs, psCodUsuario, pdFecha, "", "'" & psCodAgencia & "'", pbFecha)
        Set odCredDoc = Nothing
        
        For I = 1 To nTotal
            RaiseEvent ProgresoBarra(I, "Consolidado por Analista V 2.0", "Procesando...")
        Next I
        RaiseEvent FinalizaBarra
    Else
        Exit Function
    End If
End Function


Sub ConfiguararRsAnalista()
    Set rsAnalista = New ADODB.Recordset
    
    With rsAnalista.Fields
        .Append "FechaProceso", adDate
        .Append "Usuario", adChar, 4
        .Append "CodAna", adVarChar, 10
        .Append "NumCreVig", adInteger
        .Append "KDesemVig", adInteger
        .Append "SaldoKVig", adCurrency
        .Append "SaldoPromVig", adCurrency
        .Append "NumCreNewMes", adInteger
        .Append "KDesemNewMes", adCurrency
        .Append "NumCreRepreMes", adInteger
        .Append "KDesemRepreMes", adCurrency
        .Append "NumCreVen1a15", adInteger
        .Append "K1a15", adCurrency
        .Append "Por1a15", adDecimal
        .Append "NumCreVen16a30", adInteger
        .Append "K16a30", adCurrency
        .Append "Por16a30", adDecimal
        .Append "NumCreVen31aN", adInteger
        .Append "K31aN", adCurrency
        .Append "Por31aN", adDecimal
        .Append "NumCreJud", adInteger
        .Append "KJudicial", adCurrency
        .Append "PorJud", adDecimal
        .Append "NumJudVen", adInteger
        .Append "KJudVen", adCurrency
        .Append "PorJudVen", adDecimal
    End With
    rsAnalista.Open
End Sub

Sub ConfigurarRsAnaProducto()
    Set rsAnaProducto = New ADODB.Recordset
    
    With rsAnaProducto.Fields
        .Append "FechaProceso", adDate
        .Append "Usuario", adChar, 4
        .Append "CodAna", adVarChar, 10
        .Append "NumCreVig", adInteger
        .Append "KDesemVig", adInteger
        .Append "SaldoKVig", adCurrency
        .Append "SaldoPromVig", adCurrency
        .Append "NumCreNewMes", adInteger
        .Append "KDesemNewMes", adCurrency
        .Append "NumCreRepreMes", adInteger
        .Append "KDesemRepreMes", adCurrency
        .Append "NumCreVen1a15", adInteger
        .Append "K1a15", adCurrency
        .Append "Por1a15", adDecimal
        .Append "NumCreVen16a30", adInteger
        .Append "K16a30", adCurrency
        .Append "Por16a30", adDecimal
        .Append "NumCreVen31aN", adInteger
        .Append "K31aN", adCurrency
        .Append "Por31aN", adDecimal
        .Append "NumCreJud", adInteger
        .Append "KJudicial", adCurrency
        .Append "PorJud", adDecimal
        .Append "NumJudVen", adInteger
        .Append "KJudVen", adCurrency
        .Append "PorJudVen", adDecimal
        .Append "cProducto", adVarChar, 3
    End With
    rsAnaProducto.Open
End Sub

Public Function ImprimeMoraXAnalitaNew(ByVal pMatAgencias As Variant, ByVal psFecIni As String, _
ByVal psMoneda As String, ByVal psCodUser As String, ByVal pdFecSis As Date, _
ByVal psNomAge As String, ByVal pMatCond As Variant, ByVal pMatProd As Variant, ByVal pMatAnalistas As Variant, _
ByVal pnDiasIni As Integer, ByVal pnDiasFin As Integer, ByVal pnMontoIni As Double, ByVal pnMontoFin As Double, _
Optional psUbiGeoCod As String = "", _
Optional ByVal psNomCmac As String) As String


    Dim rsCabecera As ADODB.Recordset
    Dim rsDetalle As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim rsGarant As ADODB.Recordset
    
    Dim sAgencias As String
    Dim sCodUser As String
    Dim sProductos As String
    Dim sAnalistas As String
    Dim sCondicion As String
    
    Dim I As Integer
    Dim oCredDoc As DCredDoc
    
    Dim sCadImp As String
    
    Dim nTotal As Integer
    Dim sUserOld As String
    
    Dim sTelefono As String
    Dim sTelefono1 As String
    Dim sTelefono2 As String
    Dim sDireccionTrabajo As String
    Dim sConvenio As String
    Dim nNumLineas    As Integer
    sCadImp = ""
    
    
    Dim totSaldoKs As Double, totcobrars As Double, J As Integer 'peac 20071009
    Dim totSaldoKd As Double, totcobrard As Double 'peac 20071009
    
    '*** peac 20071023
    Dim totalCapiS As Double, totCapiS As Double
    Dim totalInteS As Double, totInteS As Double
    Dim totalMoraS As Double, totMoraS As Double
    Dim totalgastS As Double, totgastS As Double
    'ALPA
    '07/03/2008
    Dim totalCapiUs As Double, totCapiUs As Double
    Dim totalInteUs As Double, totInteUs As Double
    Dim totalMoraUs As Double, totMoraUs As Double
    Dim totalgastUs As Double, totgastUs As Double
    
    Dim totalCapiUd As Double, totCapiUd As Double
    Dim totalInteUd As Double, totInteUd As Double
    Dim totalMoraUd As Double, totMoraUd As Double
    Dim totalgastUd As Double, totgastUd As Double
        
    Dim Contador As Integer
    Dim CrediPorAna As Integer
    
    CrediPorAna = 0
    
    Dim totSaldoKUs As Double, totcobrarUs As Double
    Dim totSaldoKUd As Double, totcobrarUd As Double
    Dim intVerGarante As Integer
    Dim intVerCuota As Integer
    Dim cCodCreTemp As String
    intVerGarante = 0
    intVerCuota = 0
    Dim sCadAGara As String
    Dim sCadCuota As String
    Dim sSubTotal As String
    Dim sNombGara As String
    Dim intcuota As Integer
    'Fin ALPA
    Dim totalCapiD As Double, totCapiD As Double
    Dim totalInteD As Double, totInteD As Double
    Dim totalMoraD As Double, totMoraD As Double
    Dim totalgastD As Double, totgastD As Double
    
    '******************
    
    For I = 0 To UBound(pMatAgencias) - 1
        If I = 0 Then
            sAgencias = "" & pMatAgencias(I) & ""
        Else
            sAgencias = sAgencias & "," & pMatAgencias(I) & ""
        End If
    Next I

    For I = 0 To UBound(pMatProd) - 1
        If I = 0 Then
            sProductos = "" & pMatProd(I) & ""
        Else
            sProductos = sProductos & "," & pMatProd(I) & ""
        End If
    Next I

    For I = 0 To UBound(pMatCond) - 1
        If I = 0 Then
            sCondicion = "" & pMatCond(I) & ""
        Else
            sCondicion = sCondicion & "," & pMatCond(I) & ""
        End If
    Next I
    
    For I = 0 To UBound(pMatAnalistas) - 1
        If I = 0 Then
            sAnalistas = "" & pMatAnalistas(I) & ""
        Else
            sAnalistas = sAnalistas & "," & pMatAnalistas(I) & ""
        End If
    Next I
    
    'Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "MORA POR ANALISTA AL " & Format(psFecIni, "dd/mm/yyyy"), psNomCmac, , True)
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "MORA Y MORA PROYECTADA POR ANALISTA AL " & Format(psFecIni, "dd/mm/yyyy"), psNomCmac, , True)
    sCadImp = sCadImp & Chr(10)
    
    sCadImp = sCadImp & " RANGO DE MORA: " & pnDiasIni & " A " & pnDiasFin & " DIAS"
    
    sCadImp = sCadImp & Chr(10)
    
    Set oCredDoc = New DCredDoc
    Set rsCabecera = oCredDoc.ListaMoraXAnalistaCabecera(sAnalistas, sProductos, sAgencias, sCondicion, psMoneda, psFecIni, pnDiasIni, pnDiasFin, pnMontoIni, pnMontoFin, Right(psUbiGeoCod, 15))
    Set oCredDoc = Nothing
   ' Set rsCabecera = R
    
    If rsCabecera.RecordCount > 0 Then
        nTotal = rsCabecera.RecordCount
        RaiseEvent IniciaBarra(nTotal)
    Else
        Set rsCabecera = Nothing
        Exit Function
    End If
    
    If Not rsCabecera.EOF And Not rsCabecera.BOF Then
        rsCabecera.MoveFirst
        sUserOld = rsCabecera!cUser
    End If
    
    sCadImp = sCadImp & " Analista:" & rsCabecera!cUser & Chr(10)
    sCadImp = sCadImp & String(160, "-") & Chr(10)
    'sCadImp = sCadImp & ImpreFormat("CREDITO/CLIENTE", 38) & ImpreFormat("DIRECCION", 44) & ImpreFormat("PREST", 10) & Space(9) & ImpreFormat("CUO APROB", 13) & ImpreFormat("SALDO.K", 13) & Space(8) & ImpreFormat("MONTO.COBR", 13) & Chr(10)
    sCadImp = sCadImp & ImpreFormat("CREDITO/CLIENTE", 26) & ImpreFormat("F. FALLEC.", 10) & ImpreFormat("DIRECCION", 44) & ImpreFormat("PREST", 10) & Space(9) & ImpreFormat("CUOTAS", 13) & ImpreFormat("SALDO", 9) & Space(3) & ImpreFormat("PROMEDIO", 8) & Space(2) & ImpreFormat("MONTO", 10) & lnSaltoLinDoc
               sCadImp = sCadImp & Space(38) & Space(44) & Space(24) & ImpreFormat("APROBADAS", 13) & ImpreFormat("CAPITAL", 9) & Space(3) & ImpreFormat("ATRASO", 8) & Space(2) & ImpreFormat("COBRADO", 10) & lnSaltoLinDoc
    sCadImp = sCadImp & String(160, "-") & Chr(10)
    nNumLineas = 13
     rsCabecera.MoveFirst
     J = 0 'peac 20071009
     
    totalCapiS = 0: totalInteS = 0: totalMoraS = 0: totalgastS = 0
    totalCapiD = 0: totalInteD = 0: totalMoraD = 0: totalgastD = 0
          
     Do Until rsCabecera.EOF
            J = J + 1 'peac 20071009
            I = I + 1
            'RaiseEvent ProgresoBarra(i, "MORA DE ANALISTA", "PROCESANDO")
            RaiseEvent ProgresoBarra(I, "MORA Y MORA PROYECTADA POR ANALISTA", "PROCESANDO")
            totCapiS = 0: totInteS = 0: totMoraS = 0: totgastS = 0
            totCapiD = 0: totInteD = 0: totMoraD = 0: totgastD = 0
                        
            If sUserOld <> rsCabecera!cUser Then
               'cambio de analista
               '--**************** DATOS ADICIONALES****************************
'                        Set oCredDoc = New DCredDoc
'                        sDireccionTrabajo = oCredDoc.ObtenerDireccTrabajoTitular(rsCabecera!cCtaCod)
'                        sConvenio = oCredDoc.ObtenerConvenio(rsCabecera!cCtaCod)
'                        Set rs = oCredDoc.ObtenerTelefonosTitular(rsCabecera!cCtaCod)
'                        Set oCredDoc = Nothing
'                        'sDireccionTrabajo = sDireccionTrabajo
'                        If Not rs.EOF And Not rs.BOF Then
'                            If Not IsNull(rs!cPersTelefono) Then
'                                'bTelefono = True
'                                If Len(Trim(rs!cPersTelefono)) > 0 Then
'                                    sTelefono = "TELEFONO: " & Trim(rs!cPersTelefono)
'                                Else
'                                    sTelefono = "TELEFONO: "
'                                End If
'                            ElseIf Not IsNull(Trim(rs!cPersTelefono2)) Then
''                                bTelefono = True
'                                If Len(rs!cPersTelefono2) > 0 Then
'                                    sTelefono = "Telefono: " & Trim(rs!cPersTelefono2)
'                                End If
'                            Else
'                                sTelefono = "Telefono:"
'                            End If
'                        Else
'                            sTelefono = "Telefono:"
'                        End If
'                        Set rs = Nothing
'               '***************************************************************
                Contador = Contador + 1
                If Contador > 1 Then
                    CrediPorAna = CrediPorAna + 1
                End If
                    sCadCuota = sCadCuota & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & sCadAGara & sCadCuota
                    sCadCuota = ""
                    sCadAGara = ""
                    sCadImp = sCadImp & "TOTAL NUM. CREDITOS : " & ImpreFormat(CrediPorAna, 4, 0) & lnSaltoLinDoc
                    sCadImp = sCadImp & "TOTAL MONTOS SOLES  :" & lnSaltoLinDoc
                    sCadImp = sCadImp & "CAPITAL :" & ImpreFormat(totCapiUs, 12, 2, True) & " INTERES :" & ImpreFormat(totInteUs, 12, 2, True) & " MORA :" & ImpreFormat(totMoraUs, 12, 2, True) & " GASTOS :" & ImpreFormat(totgastUs, 12, 2, True) & " TOTAL :" & ImpreFormat(totcobrarUs, 12, 2, True) & " SALDO K. :" & ImpreFormat(totSaldoKUs, 12, 2, True) & lnSaltoLinDoc
                    sCadImp = sCadImp & "TOTAL MONTOS DOLARES: " & lnSaltoLinDoc
                    sCadImp = sCadImp & "CAPITAL :" & ImpreFormat(totCapiUd, 12, 2, True) & " INTERES :" & ImpreFormat(totInteUd, 12, 2, True) & " MORA :" & ImpreFormat(totMoraUd, 12, 2, True) & " GASTOS :" & ImpreFormat(totgastUd, 12, 2, True) & " TOTAL :" & ImpreFormat(totcobrarUd, 12, 2, True) & " SALDO K. :" & ImpreFormat(totSaldoKUd, 12, 2, True) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    
                    totCapiUs = 0
                    totInteUs = 0
                    totMoraUs = 0
                    totgastUs = 0
                    
                    totCapiUd = 0
                    totInteUd = 0
                    totMoraUd = 0
                    totgastUd = 0
                    CrediPorAna = 0
                    
                    totSaldoKUd = 0
                    totcobrarUd = 0
                    totSaldoKUs = 0
                    totcobrarUs = 0
                    
                'Por Analista
                If Mid(rsCabecera!cCtaCod, 9, 1) = "1" Then
                    totCapiUs = totCapiUs + rsCabecera!Capital
                    totInteUs = totInteUs + rsCabecera!Interes
                    totMoraUs = totMoraUs + rsCabecera!Mora
                    totgastUs = totgastUs + rsCabecera!Gastos
                    totSaldoKUs = totSaldoKUs + rsCabecera!SaldoK
                    totcobrarUs = totcobrarUs + rsCabecera!MontoCobrar
                Else
                    totCapiUd = totCapiUd + rsCabecera!Capital
                    totInteUd = totInteUd + rsCabecera!Interes
                    totMoraUd = totMoraUd + rsCabecera!Mora
                    totgastUd = totgastUd + rsCabecera!Gastos
                    
                    totSaldoKUd = totSaldoKUd + rsCabecera!SaldoK
                    totcobrarUd = totcobrarUd + rsCabecera!MontoCobrar
                End If

               sUserOld = rsCabecera!cUser
               SaltaLinea sCadImp, nNumLineas
               sCadImp = sCadImp & " Analista:" & rsCabecera!cUser & lnSaltoLinDoc
               'SaltaLinea sCadImp, nNumLineas
               sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
               sCadImp = sCadImp & ImpreFormat("CREDITO/CLIENTE", 38) & ImpreFormat("DIRECCION", 44) & ImpreFormat("PREST", 10) & Space(9) & ImpreFormat("CUOTAS", 13) & ImpreFormat("SALDO", 9) & Space(3) & ImpreFormat("PROMEDIO", 8) & Space(2) & ImpreFormat("MONTO", 10) & lnSaltoLinDoc
               sCadImp = sCadImp & Space(38) & Space(44) & Space(24) & ImpreFormat("APROBADAS", 13) & ImpreFormat("CAPITAL", 9) & Space(3) & ImpreFormat("ATRASO", 8) & Space(2) & ImpreFormat("COBRADO", 10) & lnSaltoLinDoc
               sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
               'SaltaLinea sCadImp, nNumLineas
'               sCadImp = sCadImp & ImpreFormat("CREDITO/CLIENTE", 38) & ImpreFormat("DIRECCION", 44) & ImpreFormat("PREST", 10) & Space(9) & ImpreFormat("CUO APROB", 13) & ImpreFormat("SALDO.K", 13) & Space(8) & ImpreFormat("TASA MORAT", 13) & Space(8) & ImpreFormat("MONTO COBR", 13) & lnSaltoLinDoc
'               'SaltaLinea sCadImp, nNumLineas
'               sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
'               With rsCabecera
'                 'SaltaLinea sCadImp, nNumLineas
'                 sCadImp = sCadImp & ImpreFormat(!Titular, 40) & "DOMIC: " & ImpreFormat(!Direccion, 40) & ImpreFormat(Format(!KDesembolsado, "#0.00"), 10) & ImpreFormat(!NroCuotaApr, 13, 0) & Space(9) & ImpreFormat(Format(!SaldoK, "#0.00"), 13) & ImpreFormat(!TasaMor, 13) & Space(8) & ImpreFormat(!montocobrar, 13) & lnSaltoLinDoc
'                 'SaltaLinea sCadImp, nNumLineas
'                 sCadImp = sCadImp & ImpreFormat(Mid(!cCtaCod, 1, 3), 3) & "-" & ImpreFormat(Mid(!cCtaCod, 4, 2), 2) & "-" & ImpreFormat(Mid(!cCtaCod, 6, 3), 3) & "-" & ImpreFormat(Mid(!cCtaCod, 9, 1), 1) & "-" & ImpreFormat(Mid(!cCtaCod, 10, 9), 10) & Space(9) & "TRABA: " & ImpreFormat(sDireccionTrabajo, 40) & "  " & sTelefono & "  " & Trim(sConvenio) & " CRED:" & !cRFA & "    VIGENCIA: " & ImpreFormat(!dVigencia, 10, 0) & lnSaltoLinDoc
'                 sTelefono = "Telefono: "
'               End With
'             'SaltaLinea sCadImp, nNumLineas
'
'             'cargando las garantias
'                sCadImp = sCadImp & ImpreFormat("GARANTE", 38) & ImpreFormat("DIRECC", 40) & lnSaltoLinDoc
'                Set oCredDoc = New DCredDoc
'                Set rsGarant = oCredDoc.ObtenerGarantes(rsCabecera!cCtaCod)
'                Do Until rsGarant.EOF
'                    sCadImp = sCadImp & ImpreFormat(rsGarant!cPersNombre, 38) & ImpreFormat(rsGarant!cPersDireccDomicilio, 40) & lnSaltoLinDoc
'                    rsGarant.MoveNext
'                Loop
'                Set rsGarant = Nothing
'
'                sCadImp = sCadImp & Space(66) & ImpreFormat("CUOTA", 15) & ImpreFormat("FEC.VENC", 15) & ImpreFormat("S.CUOTA", 15) & ImpreFormat("ATR.CUOTA", 15) & lnSaltoLinDoc
'                        Set oCredDoc = New DCredDoc
'                        Set rsDetalle = oCredDoc.ListaMoraXAnalistaDetalle(rsCabecera!cCtaCod, psFecIni)
'                        Set oCredDoc = Nothing
'                '        SaltaLinea sCadImp, nNumLineas
'                        'sCadImp = sCadImp & Space(66) & "*************** CUOTAS PENDIENTES*************" & Chr(10)
'                        'sCadImp = sCadImp & Chr(10)
'
'                        Do Until rsDetalle.EOF
'                 '           SaltaLinea sCadImp, nNumLineas
'                            sCadImp = sCadImp & Space(64) & ImpreFormat(rsDetalle!nCuota, 6, 0) & Space(10) & ImpreFormat(rsDetalle!dFecha, 10) & Space(9) & ImpreFormat(Format(rsDetalle!SaldoCuota, "#0.00"), 10) & ImpreFormat(rsDetalle!AtrasoCuota, 10, 0) & lnSaltoLinDoc
'
'                            If Mid(rsCabecera!cCtaCod, 9, 1) = "1" Then
'                                totCapiS = totCapiS + rsDetalle!capital
'                                totInteS = totInteS + rsDetalle!Interes
'                                totMoraS = totMoraS + rsDetalle!Mora
'                                totgastS = totgastS + rsDetalle!Gastos
'                            Else
'                                totCapiD = totCapiD + rsDetalle!capital
'                                totInteD = totInteD + rsDetalle!Interes
'                                totMoraD = totMoraD + rsDetalle!Mora
'                                totgastD = totgastD + rsDetalle!Gastos
'                            End If
                            'Total
                            If Mid(rsCabecera!cCtaCod, 9, 1) = "1" Then
                                totCapiS = totCapiS + rsCabecera!Capital
                                totInteS = totInteS + rsCabecera!Interes
                                totMoraS = totMoraS + rsCabecera!Mora
                                totgastS = totgastS + rsCabecera!Gastos
                            Else
                                totCapiD = totCapiD + rsCabecera!Capital
                                totInteD = totInteD + rsCabecera!Interes
                                totMoraD = totMoraD + rsCabecera!Mora
                                totgastD = totgastD + rsCabecera!Gastos
                            End If
'
'                            rsDetalle.MoveNext
'                        Loop
'                        Set rsDetalle = Nothing
                'SaltaLinea sCadImp, nNumLineas
'               sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
            Else
''****Modificado por ALPA
''****07/03/2008
CrediPorAna = CrediPorAna + 1
'               ' es el mismo analista
'               '--**************** DATOS ADICIONALES****************************
''                        Set oCredDoc = New DCredDoc
''                        sDireccionTrabajo = oCredDoc.ObtenerDireccTrabajoTitular(rsCabecera!cCtaCod)
''                        sConvenio = oCredDoc.ObtenerConvenio(rsCabecera!cCtaCod)
''                        Set rs = oCredDoc.ObtenerTelefonosTitular(rsCabecera!cCtaCod)
''                        Set oCredDoc = Nothing
''                        'sDireccionTrabajo = "Dir.Trabaj: " & sDireccionTrabajo
''                        If Not rs.EOF And Not rs.BOF Then
''                            If Not IsNull(rs!cPersTelefono) Then
''                               ' bTelefono = True
''                                If Len(Trim(rs!cPersTelefono)) > 0 Then
''                                    sTelefono1 = Trim(rs!cPersTelefono)
''                                Else
''                                    sTelefono1 = ""
''                                End If
''                            ElseIf Not IsNull(Trim(rs!cPersTelefono2)) Then
'''                                bTelefono = True
''                                If Len(rs!cPersTelefono2) > 0 Then
''                                    sTelefono2 = Trim(rs!cPersTelefono2)
''                                End If
''                            Else
''                                sTelefono2 = ""
''                            End If
''                        Else
''                            sTelefono1 = "Telefono:" & sTelefono1 & sTelefono2
''                        End If
''                        Set rs = Nothing
''                        sTelefono1 = "Telefono:"
'               '***************************************************************
'
''               With rsCabecera
'               ' SaltaLinea sCadImp, nNumLineas
''******ALPA
''******06/02/2008
''                sCadImp = sCadImp & ImpreFormat(Trim(!Titular), 40) & "DOMIC: " & ImpreFormat(Trim(!Direccion), 40) & ImpreFormat(Format(!KDesembolsado, "#0.00"), 10) & ImpreFormat(!NroCuotaApr, 13, 0) & Space(9) & ImpreFormat(Format(!SaldoK, "#0.00"), 13) & ImpreFormat(!montocobrar, 13) & lnSaltoLinDoc
''***************
'                sCadImp = sCadImp & ImpreFormat(!Titular, 40) & "DOMIC: " & ImpreFormat(!Direccion, 40) & ImpreFormat(Format(!KDesembolsado, "#0.00"), 10) & ImpreFormat(!NroCuotaApr, 13, 0) & Space(9) & ImpreFormat(Format(!SaldoK, "#0.00"), 13) & ImpreFormat(!TasaMor, 13) & Space(8) & ImpreFormat(!montocobrar, 13) & lnSaltoLinDoc
'                'SaltaLinea sCadImp, nNumLineas
'                sCadImp = sCadImp & ImpreFormat(Mid(!cCtaCod, 1, 3), 3) & "-" & ImpreFormat(Mid(!cCtaCod, 4, 2), 2) & "-" & ImpreFormat(Mid(!cCtaCod, 6, 3), 3) & "-" & ImpreFormat(Mid(!cCtaCod, 9, 1), 1) & "-" & ImpreFormat(Mid(!cCtaCod, 10, 9), 10) & Space(9) & "TRABA: " & ImpreFormat(sDireccionTrabajo, 40) & "  " & sTelefono & "  " & Trim(sConvenio) & " CRED: " & !cRFA & "    VIGENCIA: " & ImpreFormat(!dVigencia, 10, 0) & lnSaltoLinDoc
'                sTelefono = "Telefono: "
'                '********************************
'                'cargando las garantias
'                sCadImp = sCadImp & ImpreFormat("GARANTE", 38) & ImpreFormat("DIRECC", 40) & lnSaltoLinDoc
'                '****Modificado por ALPA
'                '****07/03/2008
''                Set oCredDoc = New DCredDoc
''                Set rsGarant = oCredDoc.ObtenerGarantes(!cCtaCod)
''                Do Until rsGarant.EOF
''                    sCadImp = sCadImp & ImpreFormat(rsGarant!cPersNombre, 38) & ImpreFormat(rsGarant!cPersDireccDomicilio, 40) & lnSaltoLinDoc
''                    rsGarant.MoveNext
''                Loop
''                Set rsGarant = Nothing
'                If !NombreGarante <> "" And !DireccionGarante <> "" And !telefonoGarante <> "" Then
'                CadImp = sCadImp & ImpreFormat(!NombreGarante, 38) & ImpreFormat(!DireccionGarante, 40) & lnSaltoLinDoc
'                End If
'
'               End With
'               'SaltaLinea sCadImp, nNumLineas
'
'               sCadImp = sCadImp & Space(6) & ImpreFormat("CUOTA", 15) & ImpreFormat("FEC.VENC", 15) & ImpreFormat("CAPITAL", 15) & ImpreFormat("INTERES", 15) & ImpreFormat("MORA", 15) & ImpreFormat("GASTOS", 15) & ImpreFormat("S.CUOTA", 15) & ImpreFormat("ATR.CUOTA", 15) & lnSaltoLinDoc
''                        Set oCredDoc = New DCredDoc
'                        Set rsDetalle = oCredDoc.ListaMoraXAnalistaDetalle(rsCabecera!cCtaCod, psFecIni)
'                        Set oCredDoc = Nothing
                        'sCadImp = sCadImp & Space(66) & "*************** CUOTAS PENDIENTES*************" & Chr(10)
                        'sCadImp = sCadImp & Chr(10)
'                        Do Until rsDetalle.EOF
'                         '   SaltaLinea sCadImp, nNumLineas
'                           ' sCadImp = sCadImp & Space(4) & ImpreFormat(rsDetalle!nCuota, 6, 0) & Space(10) & ImpreFormat(rsDetalle!dFecha, 10) & Space(5) & ImpreFormat(Format(rsDetalle!capital, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsDetalle!Interes, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsDetalle!Mora, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsDetalle!Gastos, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsDetalle!SaldoCuota, "#0.00"), 10) & ImpreFormat(rsDetalle!AtrasoCuota, 10, 0) & lnSaltoLinDoc
'
'                            If Mid(rsCabecera!cCtaCod, 9, 1) = "1" Then
'                                totCapiS = totCapiS + rsDetalle!capital
'                                totInteS = totInteS + rsDetalle!Interes
'                                totMoraS = totMoraS + rsDetalle!Mora
'                                totgastS = totgastS + rsDetalle!Gastos
'                            Else
'                                totCapiD = totCapiD + rsDetalle!capital
'                                totInteD = totInteD + rsDetalle!Interes
'                                totMoraD = totMoraD + rsDetalle!Mora
'                                totgastD = totgastD + rsDetalle!Gastos
'                            End If
'
'                            rsDetalle.MoveNext
'                        Loop
'                        Set rsDetalle = Nothing
                        '***Por ALPA
'                        sCadImp = sCadImp & Space(4) & ImpreFormat(rsCabecera!nCuota, 6, 0) & Space(10) & ImpreFormat(rsCabecera!dFecha, 10) & Space(5) & ImpreFormat(Format(rsCabecera!Capital, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!Interes, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!Mora, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!Gastos, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!SaldoCuota, "#0.00"), 10) & ImpreFormat(rsCabecera!AtrazoCuota, 10, 0) & lnSaltoLinDoc
                            
                            If Mid(rsCabecera!cCtaCod, 9, 1) = "1" Then
                                totCapiS = totCapiS + rsCabecera!Capital
                                totInteS = totInteS + rsCabecera!Interes
                                totMoraS = totMoraS + rsCabecera!Mora
                                totgastS = totgastS + rsCabecera!Gastos
                            Else
                                totCapiD = totCapiD + rsCabecera!Capital
                                totInteD = totInteD + rsCabecera!Interes
                                totMoraD = totMoraD + rsCabecera!Mora
                                totgastD = totgastD + rsCabecera!Gastos
                            End If
                            'Por Analista
                            If Mid(rsCabecera!cCtaCod, 9, 1) = "1" Then
                                totCapiUs = totCapiUs + rsCabecera!Capital
                                totInteUs = totInteUs + rsCabecera!Interes
                                totMoraUs = totMoraUs + rsCabecera!Mora
                                totgastUs = totgastUs + rsCabecera!Gastos
                                
                                totSaldoKUs = totSaldoKUs + rsCabecera!SaldoK
                                totcobrarUs = totcobrarUs + rsCabecera!MontoCobrar
                            Else
                                totCapiUd = totCapiUd + rsCabecera!Capital
                                totInteUd = totInteUd + rsCabecera!Interes
                                totMoraUd = totMoraUd + rsCabecera!Mora
                                totgastUd = totgastUd + rsCabecera!Gastos
                                
                                totSaldoKUd = totSaldoKUd + rsCabecera!SaldoK
                                totcobrarUd = totcobrarUd + rsCabecera!MontoCobrar
                            End If
               'SaltaLinea sCadImp, nNumLineas
'               sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
               
            End If
            '***************ALPA
            sTelefono = "Telefono: " & rsCabecera!Telefonotitular
            With rsCabecera
            If cCodCreTemp <> rsCabecera!cCtaCod Then
              sCadCuota = sCadCuota & String(160, "-") & lnSaltoLinDoc
              sCadImp = sCadImp & sCadAGara & sCadCuota & sSubTotal
              sCadCuota = ""
              sSubTotal = ""
              sCadAGara = ""
              sCadImp = sCadImp & ImpreFormat(!Titular, 30) & ImpreFormat(IIf(IsNull(!dPersFallec), Space(11), IIf(!dPersFallec = "", "", Format(!dPersFallec, "dd/mm/yyyy"))), 11) & "DOMIC: " & ImpreFormat(!Direccion, 40) & ImpreFormat(Format(!KDesembolsado, "#0.00"), 10) & ImpreFormat(!NroCuotaApr, 13, 0) & Space(4) & ImpreFormat(Format(!SaldoK, "#0.00"), 12) & ImpreFormat(!PromDAtra, 8) & Space(2) & ImpreFormat(!MontoCobrar, 10) & lnSaltoLinDoc
              'sCadImp = sCadImp & ImpreFormat(!Titular, 30) & "DOMIC: " & ImpreFormat(!DIRECCION, 40) & ImpreFormat(Format(!KDesembolsado, "#0.00"), 10) & ImpreFormat(!NroCuotaApr, 13, 0) & Space(4) & ImpreFormat(Format(!SaldoK, "#0.00"), 12) & ImpreFormat(!PromDAtra, 8) & Space(2) & ImpreFormat(!MontoCobrar, 10) & lnSaltoLinDoc
                'SaltaLinea sCadImp, nNumLineas
                sCadImp = sCadImp & ImpreFormat(Mid(!cCtaCod, 1, 3), 3) & "-" & ImpreFormat(Mid(!cCtaCod, 4, 2), 2) & "-" & ImpreFormat(Mid(!cCtaCod, 6, 3), 3) & "-" & ImpreFormat(Mid(!cCtaCod, 9, 1), 1) & "-" & ImpreFormat(Mid(!cCtaCod, 10, 9), 10) & Space(9) & "TRABA: " & ImpreFormat(!DireccionTrabajo, 40) & "  " & ImpreFormat(sTelefono, 25) & "  " & Trim(sConvenio) & " CRED: " & !cRFA & "    VIGENCIA: " & ImpreFormat(!dVigencia, 10, 0) & lnSaltoLinDoc
            End If
                sTelefono = "Telefono: "
                'cargando las garantias
                If !NombreGarante <> "" And !DireccionGarante <> "" And !telefonoGarante <> "" Then
                    If cCodCreTemp <> rsCabecera!cCtaCod Then
                        sCadAGara = sCadAGara & ImpreFormat("GARANTE", 38) & ImpreFormat("DIRECC", 40) & ImpreFormat("TELEFONO", 40) & lnSaltoLinDoc
                        sCadAGara = sCadAGara & ImpreFormat(!NombreGarante, 38) & ImpreFormat(!DireccionGarante, 40) & ImpreFormat(!telefonoGarante, 40) & lnSaltoLinDoc
                    Else
                        If sNombGara = !NombreGarante Then
                        sCadAGara = sCadAGara
                        Else
                        sCadAGara = sCadAGara & ImpreFormat(!NombreGarante, 38) & ImpreFormat(!DireccionGarante, 40) & ImpreFormat(!telefonoGarante, 40) & lnSaltoLinDoc
                        End If
                    End If
                    sNombGara = !NombreGarante
'                    If rsCabecera!CuotasMora > 1 Then 'And rsCabecera!CuotasMora = intVerCuota - 1
'                    sCadAGara = sCadAGara & ImpreFormat(!NombreGarante, 38) & ImpreFormat(!DireccionGarante, 40) & ImpreFormat(!telefonoGarante, 40) & lnSaltoLinDoc
'                    Else
'                    sCadImp = sCadImp & ImpreFormat(!NombreGarante, 38) & ImpreFormat(!DireccionGarante, 40) & ImpreFormat(!telefonoGarante, 40) & lnSaltoLinDoc
'                    End If
                End If
               End With
               'SaltaLinea sCadImp, nNumLineas
               If cCodCreTemp <> rsCabecera!cCtaCod Then
               intcuota = 0
               If rsCabecera!CuotasMora > 1 Then
               sCadCuota = sCadCuota & Space(6) & ImpreFormat("CUOTA", 15) & ImpreFormat("FEC.VENC", 15) & ImpreFormat("CAPITAL", 15) & ImpreFormat("INTERES", 15) & ImpreFormat("MORA", 15) & ImpreFormat("GASTOS", 15) & ImpreFormat("S.CUOTA", 10) & ImpreFormat("ATR.CUOTA", 10) & ImpreFormat("TAS.MOR", 10) & ImpreFormat("DIA MAX MORA", 10) & lnSaltoLinDoc '
               sCadCuota = sCadCuota & Space(4) & ImpreFormat(rsCabecera!nCuota, 6, 0) & Space(10) & ImpreFormat(rsCabecera!dFecha, 10) & Space(5) & ImpreFormat(Format(rsCabecera!Capital, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!Interes, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!Mora, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!Gastos, "#0.00"), 10) & Space(2) & ImpreFormat(Format(rsCabecera!SaldoCuota, "#0.00"), 10) & ImpreFormat(rsCabecera!AtrazoCuota, 10, 0) & ImpreFormat(rsCabecera!TasaMor, 10, 2) & ImpreFormat(rsCabecera!MaxDiaAtra, 10, 0) & lnSaltoLinDoc
               Else
               sCadCuota = sCadCuota & Space(6) & ImpreFormat("CUOTA", 15) & ImpreFormat("FEC.VENC", 15) & ImpreFormat("CAPITAL", 15) & ImpreFormat("INTERES", 15) & ImpreFormat("MORA", 15) & ImpreFormat("GASTOS", 15) & ImpreFormat("S.CUOTA", 10) & ImpreFormat("ATR.CUOTA", 10) & ImpreFormat("TAS.MOR", 10) & ImpreFormat("DIA MAX MORA", 10) & lnSaltoLinDoc '
               sCadCuota = sCadCuota & Space(4) & ImpreFormat(rsCabecera!nCuota, 6, 0) & Space(10) & ImpreFormat(rsCabecera!dFecha, 10) & Space(5) & ImpreFormat(Format(rsCabecera!Capital, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!Interes, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!Mora, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!Gastos, "#0.00"), 10) & Space(2) & ImpreFormat(Format(rsCabecera!SaldoCuota, "#0.00"), 10) & ImpreFormat(rsCabecera!AtrazoCuota, 10, 0) & ImpreFormat(rsCabecera!TasaMor, 10, 2) & ImpreFormat(rsCabecera!MaxDiaAtra, 10, 0) & lnSaltoLinDoc
               'sCadCuota = sCadCuota & String(160, "-") & lnSaltoLinDoc
               intVerCuota = 1
               
               End If
               Else
               'sCadCuota = sCadCuota & Space(6) & ImpreFormat("CUOTA", 15) & ImpreFormat("FEC.VENC", 15) & ImpreFormat("CAPITAL", 15) & ImpreFormat("INTERES", 15) & ImpreFormat("MORA", 15) & ImpreFormat("GASTOS", 15) & ImpreFormat("S.CUOTA", 10) & ImpreFormat("ATR.CUOTA", 10) & ImpreFormat("TAS.MOR", 10) & ImpreFormat("DIA MAX MORA", 10) & lnSaltoLinDoc '
               If intcuota <> rsCabecera!nCuota Then
               sCadCuota = sCadCuota & Space(4) & ImpreFormat(rsCabecera!nCuota, 6, 0) & Space(10) & ImpreFormat(rsCabecera!dFecha, 10) & Space(5) & ImpreFormat(Format(rsCabecera!Capital, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!Interes, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!Mora, "#0.00"), 10) & Space(5) & ImpreFormat(Format(rsCabecera!Gastos, "#0.00"), 10) & Space(2) & ImpreFormat(Format(rsCabecera!SaldoCuota, "#0.00"), 10) & ImpreFormat(rsCabecera!AtrazoCuota, 10, 0) & ImpreFormat(rsCabecera!TasaMor, 10, 2) & ImpreFormat(rsCabecera!MaxDiaAtra, 10, 0) & lnSaltoLinDoc
               intVerCuota = intVerCuota + 1
               End If
               End If
               intcuota = rsCabecera!nCuota
              'If rsCabecera!CuotasMora > 1 And rsCabecera!CuotasMora = intVerCuota Then
'               sCadImp = sCadImp & sCadCuota
                'sCadCuota = sCadCuota & String(160, "-") & lnSaltoLinDoc
'               sCadCuota = ""
               'End If
               'SaltaLinea sCadImp, nNumLineas
               
            '***************Fin ALPA
            'peac 20071009
            If Mid(rsCabecera!cCtaCod, 9, 1) = "1" Then
                totalCapiS = totalCapiS + totCapiS
                totalInteS = totalInteS + totInteS
                totalMoraS = totalMoraS + totMoraS
                totalgastS = totalgastS + totgastS
                
                totSaldoKs = totSaldoKs + rsCabecera!SaldoK
                totcobrars = totcobrars + rsCabecera!MontoCobrar
            Else
                totalCapiD = totalCapiD + totCapiD
                totalInteD = totalInteD + totInteD
                totalMoraD = totalMoraD + totMoraD
                totalgastD = totalgastD + totgastD
                                
                totSaldoKd = totSaldoKd + rsCabecera!SaldoK
                totcobrard = totcobrard + rsCabecera!MontoCobrar
            End If
        cCodCreTemp = rsCabecera!cCtaCod
        rsCabecera.MoveNext
     Loop
    sCadImp = sCadImp & sCadAGara & sCadCuota
    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & "TOTAL NUM. CREDITOS : " & ImpreFormat(CrediPorAna, 4, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & "TOTAL MONTOS SOLES  :" & lnSaltoLinDoc
    sCadImp = sCadImp & "CAPITAL :" & ImpreFormat(totCapiUs, 12, 2, True) & " INTERES :" & ImpreFormat(totInteUs, 12, 2, True) & " MORA :" & ImpreFormat(totMoraUs, 12, 2, True) & " GASTOS :" & ImpreFormat(totgastUs, 12, 2, True) & " TOTAL :" & ImpreFormat(totcobrarUs, 12, 2, True) & " SALDO K. :" & ImpreFormat(totSaldoKUs, 12, 2, True) & lnSaltoLinDoc
    sCadImp = sCadImp & "TOTAL MONTOS DOLARES: " & lnSaltoLinDoc
    sCadImp = sCadImp & "CAPITAL :" & ImpreFormat(totCapiUd, 12, 2, True) & " INTERES :" & ImpreFormat(totInteUd, 12, 2, True) & " MORA :" & ImpreFormat(totMoraUd, 12, 2, True) & " GASTOS :" & ImpreFormat(totgastUd, 12, 2, True) & " TOTAL :" & ImpreFormat(totcobrarUd, 12, 2, True) & " SALDO K. :" & ImpreFormat(totSaldoKUd, 12, 2, True) & lnSaltoLinDoc
    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
    
    'peac 20071009
    sCadImp = sCadImp & "TOTAL NUM. CREDITOS : " & ImpreFormat(J, 4, 0) & lnSaltoLinDoc
    sCadImp = sCadImp & "TOTAL MONTOS SOLES  :" & lnSaltoLinDoc
    sCadImp = sCadImp & "CAPITAL :" & ImpreFormat(totalCapiS, 12, 2, True) & " INTERES :" & ImpreFormat(totalInteS, 12, 2, True) & " MORA :" & ImpreFormat(totalMoraS, 12, 2, True) & " GASTOS :" & ImpreFormat(totalgastS, 12, 2, True) & " TOTAL :" & ImpreFormat(totcobrars, 12, 2, True) & " SALDO K. :" & ImpreFormat(totSaldoKs, 12, 2, True) & lnSaltoLinDoc
    sCadImp = sCadImp & "TOTAL MONTOS DOLARES: " & lnSaltoLinDoc
    sCadImp = sCadImp & "CAPITAL :" & ImpreFormat(totalCapiD, 12, 2, True) & " INTERES :" & ImpreFormat(totalInteD, 12, 2, True) & " MORA :" & ImpreFormat(totalMoraD, 12, 2, True) & " GASTOS :" & ImpreFormat(totalgastD, 12, 2, True) & " TOTAL :" & ImpreFormat(totcobrard, 12, 2, True) & " SALDO K. :" & ImpreFormat(totSaldoKd, 12, 2, True) & lnSaltoLinDoc
    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
     
     Set rsCabecera = Nothing
     RaiseEvent FinalizaBarra
     ImprimeMoraXAnalitaNew = sCadImp
     
End Function

Sub SaltaLinea(ByRef sCadena As String, ByRef pnumLineas As Integer)
    pnumLineas = pnumLineas + 1
    If pnumLineas > 53 Then
        sCadena = sCadena & Chr(12)
        pnumLineas = 0
    End If
End Sub

Public Function ListaDesembolsosXAgenciaXFecha(ByVal pMatAgencias As Variant, pdFecha As Date, ByVal psNomAge As String, _
ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psFecIni As String, ByVal psNomCmac As String) As String
    Dim sAgencias As String
    Dim rs As ADODB.Recordset
    Dim sCadImp As String
    Dim oCreDoc As DCredDoc
    Dim I As Integer
    Dim nTotal As Integer
    
    For I = 0 To UBound(pMatAgencias) - 1
        If I = 0 Then
            sAgencias = "'" & pMatAgencias(I) & "'"
        Else
            sAgencias = sAgencias & ",'" & pMatAgencias(I) & "'"
        End If
    Next I
    
    Set oCreDoc = New DCredDoc
    Set rs = oCreDoc.ListaDesembolsosXAgenciaXFecha(sAgencias, psFecIni)
    Set oCreDoc = Nothing
    
    If rs.RecordCount > 0 Then
        nTotal = rs.RecordCount
        RaiseEvent IniciaBarra(nTotal)
    Else
        Exit Function
    End If
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "LISTA DE DESEMBOLSOS AL " & Format(psFecIni, "dd/mm/yyyy"), psNomCmac, , True)
    sCadImp = sCadImp & String(150, "-")
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & ImpreFormat("NRO DE CREDITO", 20) & ImpreFormat("TITULAR", 25) & Space(10) & ImpreFormat("K.DESEM", 10) & ImpreFormat("MONEDA", 8) & ImpreFormat("ANALISTA", 15) & Space(15) & ImpreFormat("FECHA", 10) & ImpreFormat("USUARIO", 10) & Chr(10)
     
    Do Until rs.EOF
        I = I + 1
        RaiseEvent ProgresoBarra(I, "CREDITOS DESEMBOLSADOS", "PROCESANDO...")
        sCadImp = sCadImp & ImpreFormat(rs!cCtaCod, 20) & ImpreFormat(Trim(rs!cTitular), 35) & ImpreFormat(Format(rs!MontoDesem, "#0.00"), 10, 2) & ImpreFormat(IIf(Mid(rs!cCtaCod, 9, 1) = "1", "SOLES", "DOLARES"), 8) & ImpreFormat(Trim(rs!cAnalista), 30) & ImpreFormat(rs!dFecha, 10) & ImpreFormat(rs!cUsuario, 10) & Chr(10)
        rs.MoveNext
    Loop
    RaiseEvent FinalizaBarra
    Set rs = Nothing
    ListaDesembolsosXAgenciaXFecha = sCadImp
End Function

Public Function Reporte108310_ListaClientesAutomaticos(ByVal pMatAgencias As Variant, _
ByVal pMatAnalistas As Variant, ByVal pdFecha As Date, ByVal psNomAge As String, _
ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psNomCmac, _
ByVal psProductos As String) As String
    Dim I As Integer
    Dim rs As ADODB.Recordset
    Dim sCadImp As String
    Dim odCredDoc As DCredDoc
    Dim sAgencias As String
    Dim sAnalistas As String
    Dim sSQL As String
    Dim oConec As DConecta
    Dim sUser As String
    Dim nTotal As Integer
    Dim nNumLineas As Integer
    Dim sNomAgencia As String
    Dim sNomUsuer As String
    Dim sNomProducto As String
    Dim J As Integer
    If psProductos = "'201'" Then
        sNomProducto = " MES EMPRESARIAL"
    Else
        sNomProducto = " CONSUMO PERSONAL DIRECTO"
    End If
    sNomUsuer = ""
    Set oConec = New DConecta
    oConec.AbreConexion
    For I = 0 To UBound(pMatAgencias) - 1
        sSQL = "SELECT cAgeDescripcion FROM Agencias WHERE cAgeCod='" & pMatAgencias(I) & "'"
        Set rs = oConec.CargaRecordSet(sSQL)
        If Not rs.EOF And Not rs.BOF Then
            If Len(sNomAgencia) = 0 Then
                sNomAgencia = rs!cAgeDescripcion
            Else
                sNomAgencia = sNomAgencia & ", " & rs!cAgeDescripcion
            End If
        End If
        If I = 0 Then
            sAgencias = "'" & pMatAgencias(I) & "'"
        Else
            sAgencias = ",'" & pMatAgencias(I) & "'"
        End If
    Next I
    
    oConec.CierraConexion
    Set oConec = Nothing
'    Set oConec = New DConecta
'    oConec.AbreConexion
'
'    For I = 0 To UBound(pMatAnalistas) - 1
'        sSql = "Select RH.Cuser,Pers.cPersNombre "
'        sSql = sSql & " From RRHH RH"
'        sSql = sSql & " INNER JOIN Persona Pers ON Pers.cPersCod=RH.cPersCod"
'        sSql = sSql & " Where RH.cPersCod='" & pMatAnalistas(I) & "'"
'
'        Set Rs = oConec.CargaRecordSet(sSql)
'        If Not Rs.BOF And Not Rs.EOF Then
'            If Len(sNomUsuer) = 0 Then
'                sNomUsuer = Rs!cPersNombre
'            Else
'                sNomUsuer = sNomUsuer & ", " & Rs!cPersNombre
'            End If
'            sUser = IIf(IsNull(Rs!cUser), "", Rs!cUser)
'        End If
'        Set Rs = Nothing
'
''        If I = 0 Then
''            sAnalistas = "'" & sUser & "'"
''        Else
''            sAnalistas = sAnalistas & ",'" & sUser & "'"
''        End If
'    Next I
    
    
    
    For I = 0 To UBound(pMatAnalistas) - 1
        J = 0
        Set oConec = New DConecta
        oConec.AbreConexion
        sSQL = "Select RH.Cuser,Pers.cPersNombre "
        sSQL = sSQL & " From RRHH RH"
        sSQL = sSQL & " INNER JOIN Persona Pers ON Pers.cPersCod=RH.cPersCod"
        sSQL = sSQL & " Where RH.cPersCod='" & pMatAnalistas(I) & "'"
            
        Set rs = oConec.CargaRecordSet(sSQL)
        If Not rs.BOF And Not rs.EOF Then
            sNomUsuer = rs!cpersnombre
            sUser = IIf(IsNull(rs!cUser), "", rs!cUser)
        End If
        Set rs = Nothing
        oConec.CierraConexion
        Set oConec = Nothing
        
        Set odCredDoc = New DCredDoc
        Set rs = odCredDoc.ListaClientesAutomaticos(pdFecha, sAgencias, "'" & sUser & "'", psProductos)
        Set odCredDoc = Nothing
        
        If rs.RecordCount > 0 Then
            nTotal = rs.RecordCount
            RaiseEvent IniciaBarra(nTotal)
        Else
            sCadImp = ""
            Exit Function
        End If
        
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecha, psCodUser, "REPORTE DE CLIENTES AUTOMATICOS", psNomCmac)
        sCadImp = sCadImp & Chr(10)
        sCadImp = sCadImp & "AGENCIA: " & sNomAgencia & Chr(10)
        sCadImp = sCadImp & "ANALISTA: " & sNomUsuer & Chr(10)
        sCadImp = sCadImp & "PRODUCTO: " & sNomProducto & Chr(10)
        sCadImp = sCadImp & String(150, "-")
        sCadImp = sCadImp & Chr(10)
        ' mostrando las columnas
        sCadImp = sCadImp & ImpreFormat("Nro de Credito", 18) & ImpreFormat("Cliente", 35) & ImpreFormat("Direccion", 35) & ImpreFormat("Prestamo", 10) & ImpreFormat("Saldo.K", 10) & ImpreFormat("Calificacion", 10) & Chr(10)
        
        sCadImp = sCadImp & String(150, "-")
        sCadImp = sCadImp & Chr(10)
        Do Until rs.EOF
            If ValidaCreditoAuto(rs!cCtaCod) = True Then
                J = J + 1
                RaiseEvent ProgresoBarra(J, "Reporte de Clientes Automaticos", "Procesando...")
                sCadImp = sCadImp & ImpreFormat(rs!cCtaCod, 18) & ImpreFormat(Trim(rs!cpersnombre), 35) & ImpreFormat(Trim(rs!cPersDireccDomicilio), 35) & ImpreFormat(Format(rs!nMontoCol, "#0.00"), 10, 2) & ImpreFormat(Format(rs!nSaldo, "#0.00"), 8, 2) & Space(4) & ImpreFormat(rs!cCalificacion, 10) & lnSaltoLinDoc
            End If
            rs.MoveNext
        Loop
        Set rs = Nothing
        RaiseEvent FinalizaBarra
    Next I
    Reporte108310_ListaClientesAutomaticos = sCadImp
End Function
Function ValidaCreditoAuto(ByVal psCtaCod As String) As Boolean
    Dim oConec As DConecta
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    sSQL = "Select Count(*) as nCantidad From Producto Where cCtaCod='" & psCtaCod & "' And nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205)"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        If rs!nCantidad > 0 Then
            ValidaCreditoAuto = True
        Else
            ValidaCreditoAuto = False
        End If
    End If
    Set rs = Nothing
End Function


Public Function Repor108205_ClientesVigentes(ByVal pdFecha As Date) As String
    Dim rs As ADODB.Recordset
    Dim odCredDoc As DCredDoc
    
    Dim m_Excel As Excel.Application
    Dim oLibroExcel  As Excel.Workbook
    Dim oHojaExcel As Excel.Worksheet
    Dim oCelda  As Excel.Range
    Dim nTotal As Integer
    Dim I As Integer
    Dim nFila As Integer
    
    nTotal = 0
    
    Set odCredDoc = New DCredDoc
    Set rs = odCredDoc.Repor108205_ClientesVigentes(pdFecha)
    Set odCredDoc = Nothing
    
    If rs.RecordCount > 0 Then
        RaiseEvent IniciaBarra(rs.RecordCount)
    
    Else
        Repor108205_ClientesVigentes = ""
        Exit Function
    End If
    
    
    Set m_Excel = New Excel.Application
    'm_Excel.Visible = True
    
    Set oLibroExcel = m_Excel.Workbooks.Add
    Set oHojaExcel = oLibroExcel.Worksheets(1)
    oHojaExcel.Visible = xlSheetVisible
    
    oHojaExcel.Activate
    
    oHojaExcel.PageSetup.Zoom = 80
    oHojaExcel.PageSetup.Orientation = xlLandscape
    m_Excel.Range("A1:R1000").Font.Size = 9
    
    
    m_Excel.Selection.NumberFormat = "#,##0.00"
    
    
    'creando el encabezado del Informe
   oHojaExcel.Range("A2:D2").Merge
   oHojaExcel.Range("A2.D2").value = "LISTA DE CLIENTES DEL MES " & MonthName(Month(pdFecha))
   oHojaExcel.Range("A2:D2").Font.Italic = True
   oHojaExcel.Range("A2:D2").Font.Size = 13
   
   
   nFila = 3
  'Armando las Columnas
   nFila = nFila + 2
   
   
   With oHojaExcel
    .Range("A1").ColumnWidth = 10
    .Range("B1").ColumnWidth = 5
    .Range("C1").ColumnWidth = 5
    
    .Range("D1").ColumnWidth = 5
    .Range("E1").ColumnWidth = 5
    .Range("F1").ColumnWidth = 5
    .Range("G1").ColumnWidth = 5
    .Range("H1").ColumnWidth = 10
    
     .Range("D1").EntireColumn.NumberFormat = "0.00"
     .Range("E1").EntireColumn.NumberFormat = "0.00"
     .Range("F1").EntireColumn.NumberFormat = "dd/mm/yyyy;@"
     .Range("G1").EntireColumn.NumberFormat = "dd/mm/yyyy;@"
   End With
   
   
   With oHojaExcel
    .Cells(nFila, 1) = "NOMBRE DEL CLIENTE"
    .Cells(nFila, 1).Font.Bold = True
    
    .Cells(nFila, 2) = "EDAD"
    .Cells(nFila, 2).Font.Bold = True
    
    .Cells(nFila, 3) = "MONEDA"
    .Cells(nFila, 3).Font.Bold = True
    
    .Cells(nFila, 4) = "MONTO DESEMBOLSADO"
    .Cells(nFila, 4).Font.Bold = True
    
    .Cells(nFila, 5) = "SALDO CAPITAL"
    .Cells(nFila, 5).Font.Bold = True
    
    .Cells(nFila, 6) = "FECHA DE INICIO"
    .Cells(nFila, 6).Font.Bold = True
    
    .Cells(nFila, 7) = "FECHA DE TERMINO"
    .Cells(nFila, 7).Font.Bold = True
    
    .Cells(nFila, 8) = "NOMBRE DE LA AGENCIA"
    .Cells(nFila, 8).Font.Bold = True
   End With
   
   nFila = nFila + 1
   I = 1
    Do Until rs.EOF
        With oHojaExcel
            .Cells.EntireColumn.AutoFit
            .Cells(nFila, 1) = Trim(rs!cpersnombre)
            .Cells(nFila, 2) = Trim(rs!Edad)
            .Cells(nFila, 3) = Trim(rs!Moneda)
            .Cells(nFila, 4) = Format(rs!nMontoCol, "#0.00")
            .Cells(nFila, 5) = Format(rs!nSaldoCap, "#0.00")
            .Cells(nFila, 6) = Format(rs!dVigencia, "MM/dd/yyyy")
            .Cells(nFila, 7) = Format(rs!dVenc, "MM/dd/yyyy")
            .Cells(nFila, 8) = rs!NomAgencia
             I = I + 1
             nFila = nFila + 1
            RaiseEvent ProgresoBarra(I, "REPORTE DE CLIENTES", "CARGANDO...")
        End With
        rs.MoveNext
    Loop
    Set rs = Nothing
    RaiseEvent FinalizaBarra
    m_Excel.Visible = True
    'oLibroExcel.PrintPreview
    Repor108205_ClientesVigentes = "OK"
End Function


Public Function ListaCreditosNoDesembolsados(ByVal pMatAgencias As Variant, ByVal pdFechaInicio As Date, ByVal pdFechaFin As Date, _
ByVal pdFecSis As Date, ByVal psNomAge As String, ByVal psCodUser As String, ByVal psNomCmac As String) As String
    Dim rs As ADODB.Recordset
    Dim I As Integer
    Dim sAgencias As String
    Dim odCredDoc As DCredDoc
    Dim nTotal As Integer
    Dim nNumLineas  As Integer
    Dim sNomAgencias  As String
    Dim oConec As DConecta
    Dim sSQL As String
    Dim sCadImp As String
    
    
    For I = 0 To UBound(pMatAgencias) - 1
        If I = 0 Then
            sAgencias = "'" & pMatAgencias(I) & "'"
        Else
            sAgencias = sAgencias & ",'" & pMatAgencias(I) & "'"
        End If
        sSQL = "Select cAgeDescripcion From Agencias Where cAgeCod='" & pMatAgencias(I) & "'"
        Set oConec = New DConecta
        oConec.AbreConexion
        Set rs = oConec.CargaRecordSet(sSQL)
        oConec.CierraConexion
        Set oConec = Nothing
        If Not rs.EOF And Not rs.BOF Then
            If I = 0 Then
                sNomAgencias = rs!cAgeDescripcion
            Else
               sNomAgencias = sNomAgencias & ", " & rs!cAgeDescripcion
            End If
        End If
        Set rs = Nothing
    Next I
    
    Set odCredDoc = New DCredDoc
    Set rs = odCredDoc.ListaCreditosNoDesembolsados(pdFechaInicio, pdFechaFin, sAgencias)
    Set odCredDoc = Nothing
    
    If rs.RecordCount > 0 Then
        nTotal = rs.RecordCount
        RaiseEvent IniciaBarra(nTotal)
    Else
        sCadImp = ""
        Exit Function
    End If
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE DE CLIENTES NO DESEMBOLSADOS", psNomCmac)
    
    
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & " AGENCIA: " & sNomAgencias & Chr(10)
    sCadImp = sCadImp & " FILTRO DE FECHAS: " & pdFechaInicio & " AL " & pdFechaFin & Chr(10)
    sCadImp = sCadImp & String(150, "-")
    sCadImp = sCadImp & Chr(10)
    
    
    sCadImp = sCadImp & ImpreFormat("Numero de Credito", 20) & ImpreFormat("CLIENTE", 40) & ImpreFormat("Direccion", 40) & ImpreFormat("Monto", 8) & ImpreFormat("Fecha Ap.", 12) & ImpreFormat("Moneda", 8) & ImpreFormat("User", 6) & Chr(10)
    sCadImp = sCadImp & String(150, "-")
    sCadImp = sCadImp & Chr(10)
    
    I = 1
    Do Until rs.EOF
       I = I + 1
       sCadImp = sCadImp & ImpreFormat(rs!cCtaCod, 20) & ImpreFormat(rs!cpersnombre, 40) & ImpreFormat(rs!cPersDireccDomicilio, 40) & ImpreFormat(Format(rs!nMonto, "#0.00"), 8, 2) & ImpreFormat(rs!dPrdEstado, 12) & ImpreFormat(rs!Moneda, 8) & ImpreFormat(rs!cUser, 8) & lnSaltoLinDoc
       RaiseEvent ProgresoBarra(I, " Creditos no desembolsados", "Loading...")
       'scadImp = scadImp & impreformat(rs!cCtaCod,20) & impreformat(rs!cPersNombre,30) & impreformat(rs!cPersDireccDomicilio,30) &
       rs.MoveNext
    Loop
    Set rs = Nothing
    RaiseEvent FinalizaBarra
    ListaCreditosNoDesembolsados = sCadImp
End Function

Function ImpreKardex(ByVal prs As ADODB.Recordset, ByVal psNomAge, ByVal psFecSis As Date, ByVal psCodUser As String, _
ByVal psNomCmac As String, ByVal psCtaCod As String, ByVal psNomCliente As String) As String
    Dim sCadImp As String
    Dim I As Integer
    Dim nTotal As Integer
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, psFecSis, psCodUser, "KARDEX DEL CREDITO", psNomCmac)
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & " AGENCIA: " & psNomAge & Chr(10)
    sCadImp = sCadImp & " CLIENTE: " & psNomCliente & Chr(10)
    sCadImp = sCadImp & " NRO DE CREDITO: " & psCtaCod & Chr(10)
    sCadImp = sCadImp & String(95, "-") & Chr(10)
    sCadImp = sCadImp & Chr(10)
    
    If prs.RecordCount > 0 Then
        nTotal = prs.RecordCount
        RaiseEvent IniciaBarra(nTotal)
    Else
        sCadImp = ""
        Exit Function
    End If
    
     sCadImp = sCadImp & ImpreFormat("CUOTA", 11) & ImpreFormat("MONTO CUOTA", 10) & ImpreFormat("CAPITAL", 10) & ImpreFormat("INTERES", 15) & ImpreFormat("MORA", 10) & ImpreFormat("GASTOS", 10) & ImpreFormat("ATRASO", 10) & Chr(10)
     sCadImp = sCadImp & String(95, "-")
     
     sCadImp = sCadImp & Chr(10)
    Do Until prs.EOF
        I = I + 1
        sCadImp = sCadImp & ImpreFormat(prs!nCuota, 6) & ImpreFormat(prs!MontoCuota, 10, 2) & ImpreFormat(prs!Capital, 10, 2) & ImpreFormat(prs!Interes, 10, 2) & ImpreFormat(prs!Mora, 10) & ImpreFormat(prs!Gastos, 10) & ImpreFormat(prs!ATRASO, 10) & lnSaltoLinDoc
        RaiseEvent ProgresoBarra(I, " Kardex de Credito", "Loading...")
        prs.MoveNext
    Loop
    RaiseEvent FinalizaBarra
    ImpreKardex = sCadImp
End Function

Public Function ImpreRepor_HistorialCliente(ByVal pscPerscod As String, ByVal psNomAge As String, _
ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psNomCmac As String) As String
    Dim I As Integer
    Dim sCadImp As String
    Dim nTotal As Integer
    Dim rs As ADODB.Recordset
    Dim odCredDoc As DCredDoc
    Dim nMontoSoles As Double
    Dim nMontoDolares As Double
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "HISTORIAL DEL CLIENTE", psNomCmac)
    
    Set odCredDoc = New DCredDoc
    Set rs = odCredDoc.Repor_HistorialCliente(pscPerscod)
    Set odCredDoc = Nothing
    
    If rs.RecordCount > 0 Then
        RaiseEvent IniciaBarra(nTotal)
    Else
        ImpreRepor_HistorialCliente = ""
        Exit Function
    End If
    
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & " AGENCIA: " & psNomAge & Chr(10)
    sCadImp = sCadImp & " CLIENTE: " & rs!cNomCli & Chr(10)
    sCadImp = sCadImp & " DOC.IDENTIDAD: " & rs!cPersIDNro & Chr(10)
    sCadImp = sCadImp & String(150, "-") & Chr(10)
    sCadImp = sCadImp & Chr(10)



    sCadImp = sCadImp & ImpreFormat("CREDITO", 18) & ImpreFormat("ESTADO", 20) & ImpreFormat("DESEMBOLSO", 10) & ImpreFormat("SALDO.CAP", 10) & ImpreFormat("FEC.APROB", 10) & ImpreFormat("F.CUL/CAN", 10) & ImpreFormat("ANA", 10) & ImpreFormat("TIPO CRED.", 19) & ImpreFormat("INSTITUCION", 15) & ImpreFormat("MONEDA", 8) & Chr(10)
    sCadImp = sCadImp & String(150, "-") & Chr(10)
    
    rs.MoveFirst

    nMontoSoles = 0
    nMontoDolares = 0
    Do Until rs.EOF
        sCadImp = sCadImp & ImpreFormat(rs!cCtaCod, 18) & ImpreFormat(rs!EstadoCredito, 16) & ImpreFormat(rs!nMontoCol, 10, 2) & ImpreFormat(rs!nSaldo, 10, 2) & Space(2) & ImpreFormat(Format(rs!dPrdEstado, "dd/MM/yyyy"), 10) & ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 10) & ImpreFormat(IIf(IsNull(rs!cUser), "    ", rs!cUser), 10) & ImpreFormat(rs!TipoCredito, 19) & ImpreFormat(IIf(IsNull(rs!cNomIntitucion), "SIN CONVENIO", rs!cNomIntitucion), 15) & ImpreFormat(rs!Moneda, 10) & lnSaltoLinDoc
        If Mid(rs!cCtaCod, 9, 1) = "1" Then
            nMontoSoles = nMontoSoles + rs!nSaldo
        Else
            nMontoDolares = nMontoDolares + rs!nSaldo
        End If
        rs.MoveNext
    Loop
    
    sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Chr(10)
   
    sCadImp = sCadImp & Space(10) & "DEUDA CAPITAL EN SOLES: " & Format(nMontoSoles, "#0.00") & Space(5) & "DEUDA CAPITAL EN DOLARES: " & Format(nMontoDolares, "#0.00") & lnSaltoLinDoc
    sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc

   ImpreRepor_HistorialCliente = sCadImp
End Function

Public Function ImpreRepor_HistorialCredito(ByVal psCtaCod As String, ByVal psNomAge As String, _
ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psNomCmac As String) As String
    Dim I As Integer
    Dim sCadImp As String
    Dim nTotal As Integer
    Dim rs As ADODB.Recordset
    Dim odCredDoc As DCredDoc
    Dim nMontoCapital As Double
    Dim nMontoInteres As Double
    Dim nMontoMora As Double
    Dim nMontoGastos As Double
    
    'Dim rs As ADODB.Recordset
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "HISTORIAL DEL CREDITO", psNomCmac)
    sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & lnSaltoLinDoc
    
    Set odCredDoc = New DCredDoc
    Set rs = odCredDoc.Repor_HistorialCreditoCabecera(psCtaCod)
    Set odCredDoc = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        sCadImp = sCadImp & ImpreFormat("NRO CREDITO: ", 10) & ImpreFormat(psCtaCod, 20) & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("CLIENTE: ", 10) & ImpreFormat(rs!cpersnombre, 20) & ImpreFormat("CODIGO: ", 15) & ImpreFormat(rs!cPersCod, 15) & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("CREDITO: ", 10) & ImpreFormat(psCtaCod, 20) & ImpreFormat("VIGENCIA: ", 15) & ImpreFormat(Format(rs!dVigencia, "dd/MM/yyyy"), 15) & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("ESTADO: ", 10) & ImpreFormat(rs!cEstado, 20) & ImpreFormat("VENCIMIENTO: ", 15) & ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 15) & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("TASA: ", 10) & ImpreFormat(rs!TasaInteres & "%", 20, 2) & ImpreFormat("ANALISTA: ", 15) & ImpreFormat(rs!cUser, 15) & lnSaltoLinDoc
        'sCadImp = sCadImp & ImpreFormat("DESEMB: ", 10) & ImpreFormat(rs!nMontoCol, 20, 2) & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("DESEMB: ", 10) & Space(5) & Format(rs!nMontoCol, "#0.00") & lnSaltoLinDoc
    End If
    Set rs = Nothing
    
    Set odCredDoc = New DCredDoc
    Set rs = odCredDoc.Repo_HistorialCreditoDetalle(psCtaCod)
    Set odCredDoc = Nothing
    
    'imprimiendo las columnas del detalle
    
    sCadImp = sCadImp & Chr(10)
    sCadImp = sCadImp & String(130, "-") & Chr(10)
    sCadImp = sCadImp & ImpreFormat("NROCUOTA", 10) & ImpreFormat("FEC.VENC", 10) & ImpreFormat("FEC.VIG", 10) & ImpreFormat("CAPITAL", 8) & ImpreFormat("INTERES", 8) & Space(6) & ImpreFormat("MORA", 8) & ImpreFormat("GASTOS", 8) & ImpreFormat("ATRASO", 8) & ImpreFormat("USUARIO", 8) & lnSaltoLinDoc
    sCadImp = sCadImp & String(130, "-") & Chr(10)
    
    nMontoCapital = 0
    nMontoInteres = 0
    nMontoMora = 0
    nMontoGastos = 0
    Do Until rs.EOF
        I = I + 1
        sCadImp = sCadImp & ImpreFormat(rs!nNroCuota, 10, 0) & ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 10) & ImpreFormat(Format(rs!FechaPago, "dd/MM/yyyy"), 10) & ImpreFormat(rs!Capital, 8, 2) & ImpreFormat(rs!Interes, 8, 2) & ImpreFormat(rs!Mora, 8, 2) & ImpreFormat(rs!Gastos, 8, 2) & ImpreFormat(IIf(IsNull(rs!nDiasAtraso), "", rs!nDiasAtraso), 8) & ImpreFormat(rs!cUsuario, 8) & lnSaltoLinDoc
        If I > 1 Then
            nMontoCapital = nMontoCapital + rs!Capital
            nMontoInteres = nMontoInteres + rs!Interes
            nMontoMora = nMontoMora + rs!Mora
            nMontoGastos = nMontoGastos + rs!Gastos
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
    
    Set odCredDoc = New DCredDoc
    Set rs = odCredDoc.Repo_HistorialCredDetalleComplem(psCtaCod, pdFecSis)
    Set odCredDoc = Nothing
    
    Do Until rs.EOF
        sCadImp = sCadImp & ImpreFormat(rs!nCuota, 10, 0) & ImpreFormat(Format(rs!dVenc, "dd/MM/yyyy"), 10) & ImpreFormat(Format(IIf(IsNull(rs!FechaPago), "", rs!FechaPago), "dd/MM/yyyy"), 10) & ImpreFormat(rs!Captital, 8, 2) & ImpreFormat(rs!Interes, 8, 2) & ImpreFormat(rs!Mora, 8, 2) & ImpreFormat(rs!Gastos, 8, 2) & ImpreFormat(IIf(IsNull(rs!nDiasAtraso), "", rs!nDiasAtraso), 8) & ImpreFormat(IIf(IsNull(rs!cUsuario), "", rs!cUsuario), 8) & lnSaltoLinDoc
     rs.MoveNext
    Loop
    
    sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
    sCadImp = sCadImp & Space(35) & ImpreFormat(nMontoCapital, 8, 2) & ImpreFormat(nMontoInteres, 8, 2) & ImpreFormat(nMontoMora, 8, 2) & ImpreFormat(nMontoGastos, 8, 2) & lnSaltoLinDoc
    sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
    ImpreRepor_HistorialCredito = sCadImp
End Function

Public Function Recup_ClientesByAnalista(ByVal pMatAgencias As Variant, ByVal pMatAnalistas As Variant, ByVal psNomAge As String, _
ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psNomCmac As String) As String
    Dim I As Integer
    Dim sCadImp As String
    Dim nTotal As Integer
    Dim rs As ADODB.Recordset
    Dim odCredDoc As DCredDoc
    Dim sAgencias As String
    Dim sAnalistas As String
    Dim sUserOld As String
    
    For I = 0 To UBound(pMatAgencias) - 1
        If I = 0 Then
            sAgencias = "'" & pMatAgencias(I) & "'"
        Else
            sAgencias = ",'" & pMatAgencias(I) & "'"
        End If
    Next I
    
    For I = 0 To UBound(pMatAnalistas) - 1
        If I = 0 Then
            sAnalistas = "'" & pMatAnalistas(I) & "'"
        Else
            sAnalistas = sAnalistas & ",'" & pMatAnalistas(I) & "'"
        End If
    Next I
    
    Set odCredDoc = New DCredDoc
    Set rs = odCredDoc.Recup_ClientesByAnalista(sAnalistas, sAgencias)
    Set odCredDoc = Nothing
    
    If rs.RecordCount > 0 Then
        nTotal = rs.RecordCount
        RaiseEvent IniciaBarra(nTotal)
    Else
        Recup_ClientesByAnalista = ""
        Exit Function
    End If
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "LISTA DE CLIENTES", psNomCmac)
    sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
    
    sUserOld = rs!cUser
    sCadImp = sCadImp & "ANALISTA: " & sUserOld & lnSaltoLinDoc
    
    sCadImp = sCadImp & ImpreFormat("COD.PERS", 15) & ImpreFormat("TITULAR", 35) & ImpreFormat("DIRECCION", 35) & ImpreFormat("TELEFONO1", 15) & ImpreFormat("TELEFONO2", 15) & lnSaltoLinDoc
    sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
    
    Do Until rs.EOF
        If sUserOld = rs!cUser Then
            sCadImp = sCadImp & ImpreFormat(rs!cPersCod, 15) & ImpreFormat(rs!cpersnombre, 35) & ImpreFormat(rs!cPersDireccDomicilio, 35) & ImpreFormat(rs!cPersTelefono, 15) & ImpreFormat(rs!cPersTelefono2, 15) & lnSaltoLinDoc
        Else
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & "ANALISTA: " & rs!cUser & lnSaltoLinDoc
            sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
            
            sCadImp = sCadImp & ImpreFormat("COD.PERS", 15) & ImpreFormat("TITULAR", 35) & ImpreFormat("DIRECCION", 35) & ImpreFormat("TELEFONO1", 15) & ImpreFormat("TELEFONO2", 15) & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat(rs!cPersCod, 15) & ImpreFormat(rs!cpersnombre, 35) & ImpreFormat(rs!cPersDireccDomicilio, 35) & ImpreFormat(rs!cPersTelefono, 15) & ImpreFormat(rs!cPersTelefono2, 15) & lnSaltoLinDoc
            sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
        End If
        sUserOld = rs!cUser
        RaiseEvent ProgresoBarra(I, "ARQUEO DE CLIENTE", "LOADING...")
        rs.MoveNext
    Loop
    RaiseEvent FinalizaBarra
    Set rs = Nothing
    Recup_ClientesByAnalista = sCadImp
End Function

'Public Sub Repor_ConsolidadoProducto(ByVal pdFecha As Date, ByVal sAgencia As String, _
'ByVal pdFecSis As Date, ByVal psCodUser As String, ByVal psNomCmac As String, ByVal psNomAgencia As String)
'    Dim rs As ADODB.Recordset
'    Dim i As Integer
'    Dim odCredDoc As DCredDoc
'    Dim nTotal As Integer
'    Dim nNumLineas As Integer
'    Dim m_Excel As Excel.Application
'    Dim oLibroExcel As Excel.Workbook
'    Dim oHojaExcel As Excel.Worksheet
'    Dim oCelda As Excel.Range
'
'    Set m_Excel = New Excel.Application
'
'
'    'Armando las columnas
'    Set m_Excel = New Excel.Application
'    Set oLibroExcel = m_Excel.Workbooks.Add
'    Set oHojaExcel = oLibroExcel.Worksheets(1)
'    oHojaExcel.Visible = xlSheetVisible
'
'    oHojaExcel.Activate
'
'    oHojaExcel.PageSetup.Zoom = 80
'    oHojaExcel.PageSetup.Orientation = xlLandscape
'    m_Excel.Range("A1:R10000").Font.Size = 7
'    m_Excel.Selection.NumberFormat = "#,##0.00"
'
'    'Creando el encabezado del Informe
'       oHojaExcel.Range("2:D2").Merge
'       oHojaExcel.Range("A2.D2").value = "CONSOLIDADO DE CREDITOS "
'       oHojaExcel.Range("A2:D2").Font.Italic = True
'       oHojaExcel.Range("A2:D2").Font.Size = 13
'
'       oHojaExcel.Range("A3:D2").Merge
'       oHojaExcel.Range("A3.D2").value = "POR PRODUCTO"
'       oHojaExcel.Range("A3:D2").Font.Italic = True
'       oHojaExcel.Range("A3:D2").Font.Size = 13
'
'       oHojaExcel.Range("A3:D2").Merge
'       oHojaExcel.Range("A3.D2").value = "CAJA MUNICIPAL DE ICA S.A"
'       oHojaExcel.Range("A3:D2").Font.Italic = True
'       oHojaExcel.Range("A3:D2").Font.Size = 13
'
'       oHojaExcel.Range("A4:D2").Merge
'       oHojaExcel.Range("A4.D2").value = "Nombre de Agencia : " & psNomAgencia
'       oHojaExcel.Range("A4:D2").Font.Italic = True
'       oHojaExcel.Range("A4:D2").Font.Size = 13
'
'       oHojaExcel.Range("A5:D2").Merge
'       oHojaExcel.Range("A5.D2").value = "Usuario : " & psCodUser
'       oHojaExcel.Range("A5:D2").Font.Italic = True
'       oHojaExcel.Range("A5:D2").Font.Size = 13
'
'       oHojaExcel.Range("A6:D2").Merge
'       oHojaExcel.Range("A6.D2").value = "Fecha : " & psCodUser
'       oHojaExcel.Range("A6:D2").Font.Italic = True
'       oHojaExcel.Range("A6:D2").Font.Size = 13
'
'       oHojaExcel.Range("A7:D2").Merge
'       oHojaExcel.Range("A7.D2").value = "CARTERA TOTAL"
'       oHojaExcel.Range("A7:D2").Font.Italic = True
'       oHojaExcel.Range("A7:D2").Font.Size = 13
'
'
'       'Armando el Encabezado de Cartera Vigente
'
'       With oHojaExcel.Range("A8:E10").Select
'            .HorizontalAlignment = xlJustify
'            .VerticalAlignment = xlJustify
'            .WrapText = False
'            .Orientation = 0
'            .AddIndent = False
'            .IndentLevel = 0
'            .ShrinkToFit = False
'            .ReadingOrder = xlContext
'            .MergeCells = True
'
'            .Borders(xlDiagonalDown).LineStyle = xlNone
'            .Borders(xlDiagonalUp).LineStyle = xlNone
'
'            With .Borders(xlEdgeLeft)
'                    .LineStyle = xlContinuous
'                    .Weight = xlMedium
'                    .ColorIndex = xlAutomatic
'            End With
'
'            With .Borders(xlEdgeTop)
'                    .LineStyle = xlContinuous
'                    .Weight = xlMedium
'                    .ColorIndex = xlAutomatic
'            End With
'
'            With .Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            With .Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .Weight = xlMedium
'                .ColorIndex = xlAutomatic
'            End With
'
'            .Borders(xlInsideVertical).LineStyle = xlNone
'            .Borders(xlInsideHorizontal).LineStyle = xlNone
'       End With
'       oHojaExcel.Range("A8:E10").FormulaR1C1 = "CARTERA VIGENTE"
'
'       'Armando la Columna CodAna
'       With oHojaExcel
'            .Range("A11:A11").value = "Cod.Ana."
'            .Range("B11.B11").value = "Num.Cre"
'       End With
'End Sub
'

Function Repor_CalidadCartera(ByVal pnTipoCambio As Double, ByVal psFileName As String, ByVal pdFecha As Date, _
ByVal psRuta As String) As String
    Dim m_Excel As Excel.Application
    Dim oLibroExcel As Excel.Workbook
    Dim oHojaExcel As Excel.Worksheet
    Dim rs As ADODB.Recordset
    Dim odCredDoc As DCredDoc
    Dim nFila As Integer
    Dim I As Integer
    Dim nMontoNor3160 As Double
    Dim nMontoNor6190 As Double
    Dim nMontoReF3160 As Double
    Dim nMontoReF6190 As Double
    Dim sEstado As String
    
    Set m_Excel = New Excel.Application
    Set oLibroExcel = m_Excel.Workbooks.Open(psFileName)
    Set oHojaExcel = oLibroExcel.Worksheets(1)
    
    Set odCredDoc = New DCredDoc
    Set rs = odCredDoc.GetCalidadCartera(pnTipoCambio)
    Set odCredDoc = Nothing
    
    If rs.RecordCount = 0 Then
        Repor_CalidadCartera = "Mal"
        Exit Function
    Else
        RaiseEvent IniciaBarra(rs.RecordCount)
    End If
    
    oHojaExcel.Cells(2, 1) = "Fecha de Proceso:" & Format(pdFecha, "dd/MM/yyyy")
    I = 1
    Do Until rs.EOF
        With oHojaExcel
        
           If Trim(rs!CDescripcion) = "VIGENTE NORMAL (0)" Then
            nFila = 4
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoProv), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "VIGENTE NORMAL (1-8)" Then
            nFila = 5
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "VIGENTE NORMAL (9-30)" Then
            nFila = 6
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "VENCIDA (COM) NORMAL (16-30)" Then
            nFila = 7
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "VENCIDA NORMAL (31-60)" Then
            nFila = 8
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "VENCIDA NORMAL (61-90)" Then
            nFila = 9
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "VENCIDA NORMAL (91-120)" Then
            nFila = 10
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "VENCIDA NORMAL (121-180)" Then
            nFila = 11
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "VENCIDA NORMAL (>=181)" Then
            nFila = 12
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "REFINANCIADA (0)" Then
            nFila = 13
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "REFINANCIADA (1-8)" Then
            nFila = 14
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "REFINANCIADA (9-30)" Then
            nFila = 15
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "REFINANCIADA COM (16-30)" Then
            nFila = 16
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "REFINANCIADA (31-60)" Then
            nFila = 17
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "REFINANCIADA (61-90)" Then
            nFila = 18
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "REFINANCIADA (91-120)" Then
            nFila = 19
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "REFINANCIADA (121-180)" Then
            nFila = 20
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "REFINANCIADA (>=181)" Then
            nFila = 21
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
           If Trim(rs!CDescripcion) = "JUDICIAL" Then
            nFila = 22
            .Cells(nFila, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
            .Cells(nFila, 4) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoProv)
            .Cells(nFila, 6) = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
            .Cells(nFila, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
           End If
           
        End With
        RaiseEvent ProgresoBarra(I, "CALIDAD DE CARTERA", "PROCESANDO...")
        I = I + 1
        rs.MoveNext
    Loop
    Set rs = Nothing
    'calculando los totales
    oHojaExcel.Cells(23, 2) = "=SUM(B4:B22)"
    oHojaExcel.Cells(23, 4) = "=SUM(D4:D22)"
    oHojaExcel.Cells(23, 6) = "=SUM(F4:F22)"
    oHojaExcel.Cells(23, 8) = "=SUM(H4:H22)"
    'calculando porcentajes
    
    For I = 4 To 22
        With oHojaExcel
            .Cells(I, 3) = (.Cells(I, 2) / .Cells(23, 2)) * 100
            .Cells(I, 5) = (.Cells(I, 4) / .Cells(23, 4)) * 100
            .Cells(I, 7) = (.Cells(I, 6) / .Cells(23, 6)) * 100
            .Cells(I, 9) = (.Cells(I, 8) / .Cells(23, 8)) * 100
        End With
    Next I
    'calculanndo los totales de los porcentajes
    oHojaExcel.Cells(23, 3) = "=SUM(C4:C22)"
    oHojaExcel.Cells(23, 5) = "=SUM(E4:E22)"
    oHojaExcel.Cells(23, 7) = "=SUM(G4:G22)"
    oHojaExcel.Cells(23, 9) = "=SUM(I4:I22)"
    
    'calculando la segunda parte de la hoja
    '% de mora
    
    oHojaExcel.Cells(26, 2) = "=B23-(B13+B4)"
    oHojaExcel.Cells(27, 2) = "=(B26/B23)*100"
    
    oHojaExcel.Cells(26, 4) = "=D23-(D13+D4)"
    oHojaExcel.Cells(27, 4) = "=(D26/D23)*100"
    
    oHojaExcel.Cells(26, 6) = "=F23-(F13+F4)"
    oHojaExcel.Cells(27, 6) = "=(F26/F23)*100"
    
    oHojaExcel.Cells(26, 8) = "=H23-(H13+H4)"
    oHojaExcel.Cells(27, 8) = "=(H26/H23)*100"
    
    'Cartera atrasada
    oHojaExcel.Cells(29, 2) = "=SUM(B22,B21,B16,B12,B7)"
    oHojaExcel.Cells(30, 2) = "=(B29/B23)*100"
    
    oHojaExcel.Cells(29, 4) = "=SUM(D16:D21,D7:D12)"
    oHojaExcel.Cells(30, 4) = "=(D29/D23)*100"
    
    oHojaExcel.Cells(29, 6) = "=SUM(F16:F21,F7:F12)"
    oHojaExcel.Cells(30, 6) = "=(F29/F23)*100"
    
    oHojaExcel.Cells(29, 8) = "=SUM(H16,H17:H21,H7:H12)"
    oHojaExcel.Cells(30, 8) = "=(H29/H23)*100"
    
    'Cartera en  riesgo
    oHojaExcel.Cells(32, 2) = "=SUM(B5:B22)"
    oHojaExcel.Cells(33, 2) = "=(B32/B23)*100"
    
    oHojaExcel.Cells(32, 4) = "=SUM(D16:D21,D7:D12)"
    oHojaExcel.Cells(33, 4) = "=(D32/D23)*100"
    
    oHojaExcel.Cells(32, 6) = "=SUM(F5:F22)"
    oHojaExcel.Cells(33, 6) = "=(F32/F23)*100"
    
    oHojaExcel.Cells(32, 8) = "=SUM(H5:H22)"
    oHojaExcel.Cells(33, 8) = "=(H32/H23)*100"
    
    'Cartera en alto riesgo
    oHojaExcel.Cells(35, 2) = "=SUM(B13:B15)"
    oHojaExcel.Cells(36, 2) = "=(B35/B23)*100"
    
    oHojaExcel.Cells(35, 4) = "=SUM(D13:D15)"
    oHojaExcel.Cells(36, 4) = "=(D35/D23)*100"
    
    oHojaExcel.Cells(35, 6) = "=SUM(F13:F15)"
    oHojaExcel.Cells(36, 6) = "=(F35/F23)"
    
    oHojaExcel.Cells(35, 8) = "=SUM(H13:H15)"
    oHojaExcel.Cells(36, 8) = "=(H35/H23)*100"
    
    'obteniendo informacion de los castigados
    Set odCredDoc = New DCredDoc
    Set rs = odCredDoc.GetCalidadCarteraCastigados(pnTipoCambio)
    Set odCredDoc = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        'Cartera Castigados
        oHojaExcel.Cells(39, 2) = IIf(IsNull(rs!SaldoK), 0, rs!SaldoK)
        oHojaExcel.Cells(40, 2) = "=(B39/B23)*100"
        
        oHojaExcel.Cells(39, 4) = IIf(IsNull(rs!SaldoProv), 0, rs!SaldoProv)
        oHojaExcel.Cells(40, 4) = "=(D39/D23)*100"
        
        oHojaExcel.Cells(39, 6) = IIf(IsNull(rs!SaldoProv), 0, rs!NumCreditos)
        oHojaExcel.Cells(40, 6) = "=(F39/F23)"
        
        oHojaExcel.Cells(39, 8) = IIf(IsNull(rs!NumCliente), 0, rs!NumCliente)
        oHojaExcel.Cells(40, 8) = "=(H39/H23)*100"
   End If
    
   'adecuando los valores por los creditos de consumo
   'obteniendo los valores de los creditos vigentes normales
   sEstado = "'2020','2021','2022'"
   
   Set odCredDoc = New DCredDoc
   Set rs = odCredDoc.GetCalidadCarteraConsumo(pnTipoCambio, sEstado, 31, 60)
   If Not rs.EOF And Not rs.BOF Then
        nMontoNor3160 = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
   End If
   Set rs = odCredDoc.GetCalidadCarteraConsumo(pnTipoCambio, sEstado, 61, 90)
   If Not rs.EOF And Not rs.BOF Then
        nMontoNor6190 = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
   End If
   Set rs = Nothing
   sEstado = "'2030','2031','2032'"
   
   Set rs = odCredDoc.GetCalidadCarteraConsumo(pnTipoCambio, sEstado, 31, 60)
   If Not rs.EOF And Not rs.BOF Then
        nMontoReF3160 = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
   End If
   Set rs = odCredDoc.GetCalidadCarteraConsumo(pnTipoCambio, sEstado, 61, 90)
   If Not rs.EOF And Not rs.BOF Then
        nMontoReF6190 = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
   End If
   Set rs = Nothing
   
   With oHojaExcel
    .Cells(4, 2) = .Cells(4, 2) + nMontoNor3160
    .Cells(8, 2) = .Cells(8, 2) - nMontoNor3160
    
    .Cells(4, 2) = .Cells(4, 2) + nMontoNor6190
    .Cells(9, 2) = .Cells(9, 2) - nMontoNor6190
    
    .Cells(13, 2) = .Cells(13, 2) + nMontoReF3160
    .Cells(17, 2) = .Cells(13, 2) - nMontoReF3160
    
    .Cells(13, 2) = .Cells(13, 2) + nMontoReF6190
    .Cells(18, 2) = .Cells(13, 2) - nMontoReF6190
   End With
   
   
   
   oHojaExcel.SaveAs psRuta & "\Spooler\CalidadCartera" & Format(pdFecha, "YYYYMMDD") & ".XLS"
   oLibroExcel.Close
   m_Excel.Quit
   Set m_Excel = Nothing
   RaiseEvent FinalizaBarra
   Repor_CalidadCartera = "OK"
End Function


Public Function ImprimeMoraXAnalitaNewTelefono(ByVal pMatAgencias As Variant, ByVal psFecIni As String, _
ByVal psCodUser As String, ByVal pdFecSis As Date, _
ByVal psNomAge As String, ByVal pMatAnalistas As Variant, _
ByVal pnDiasIni As Integer, ByVal pnDiasFin As Integer, Optional ByVal psNomCmac As String) As String


    Dim rsCabecera As ADODB.Recordset
    Dim rsDetalle As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim rsGarant As ADODB.Recordset
    
    Dim sAgencias As String
    Dim sCodUser As String
    Dim sAnalistas As String
    
    
    Dim I As Integer
    Dim oCredDoc As DCredDoc
    
    Dim sCadImp As String
    
    Dim nTotal As Integer
    Dim sUserOld As String
    
    Dim sTelefono As String
    Dim sDireccionTrabajo As String
    Dim sConvenio As String
    Dim nNumLineas    As Integer
    sCadImp = ""
    
    For I = 0 To UBound(pMatAgencias) - 1
        If I = 0 Then
            sAgencias = "'" & pMatAgencias(I) & "'"
        Else
            sAgencias = sAgencias & ",'" & pMatAgencias(I) & "'"
        End If
    Next I
'
'    For i = 0 To UBound(pMatAnalistas) - 1
'        If i = 0 Then
'            sAnalistas = "'" & pMatAnalistas(i) & "'"
'        Else
'            sAnalistas = sAnalistas & ",'" & pMatAnalistas(i) & "'"
'        End If
'    Next i
     Set oCredDoc = New DCredDoc
     sAnalistas = oCredDoc.ListaAnalistasAgencia(pMatAgencias(0))
     Set oCredDoc = Nothing
     
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "MORA POR ANALISTA AL " & Format(psFecIni, "dd/mm/yyyy"), psNomCmac, , True)
    
    sCadImp = sCadImp & Chr(10)
    
    sCadImp = sCadImp & " RANGO DE MORA: " & pnDiasIni & " A " & pnDiasFin & " DIAS"
    
    sCadImp = sCadImp & Chr(10)
    
    Set oCredDoc = New DCredDoc
    Set rsCabecera = oCredDoc.ListaCreditosMoraXTelefono(sAnalistas, psFecIni, pnDiasIni, pnDiasFin, sAgencias)
    Set oCredDoc = Nothing
    
    If rsCabecera.RecordCount > 0 Then
        nTotal = rsCabecera.RecordCount
        RaiseEvent IniciaBarra(nTotal)
    Else
        Set rsCabecera = Nothing
        Exit Function
    End If
    
    If Not rsCabecera.EOF And Not rsCabecera.BOF Then
        rsCabecera.MoveFirst
        sUserOld = rsCabecera!cUser
    End If
    
    sCadImp = sCadImp & " Analista:" & rsCabecera!cUser & Chr(10)
    sCadImp = sCadImp & String(160, "-") & Chr(10)
    sCadImp = sCadImp & ImpreFormat("CREDITO/CLIENTE", 38) & ImpreFormat("DIRECCION", 44) & ImpreFormat("PREST", 10) & Space(9) & ImpreFormat("CUO APROB", 13) & ImpreFormat("SALDO.K", 13) & Space(8) & ImpreFormat("MONTO.COBR", 13) & Chr(10)
    sCadImp = sCadImp & String(160, "-") & Chr(10)
    nNumLineas = 13
     rsCabecera.MoveFirst
     Do Until rsCabecera.EOF
            I = I + 1
            RaiseEvent ProgresoBarra(I, "MORA DE ANALISTA", "PROCESANDO")
            If sUserOld <> rsCabecera!cUser Then
               'cambio de analista
               '--**************** DATOS ADICIONALES****************************
                        Set oCredDoc = New DCredDoc
                        sDireccionTrabajo = oCredDoc.ObtenerDireccTrabajoTitular(rsCabecera!cCtaCod)
                        sConvenio = oCredDoc.ObtenerConvenio(rsCabecera!cCtaCod)
                        Set rs = oCredDoc.ObtenerTelefonosTitular(rsCabecera!cCtaCod)
                        Set oCredDoc = Nothing
                        'sDireccionTrabajo = sDireccionTrabajo
                        If Not rs.EOF And Not rs.BOF Then
                            If Not IsNull(rs!cPersTelefono) Then
                                'bTelefono = True
                                If Len(Trim(rs!cPersTelefono)) > 0 Then
                                    sTelefono = "TELEFONO: " & Trim(rs!cPersTelefono)
                                Else
                                    sTelefono = "TELEFONO: "
                                End If
                            ElseIf Not IsNull(Trim(rs!cPersTelefono2)) Then
'                                bTelefono = True
                                If Len(rs!cPersTelefono2) > 0 Then
                                    sTelefono = "TELEFONO: " & Trim(rs!cPersTelefono2)
                                End If
                            Else
                                sTelefono = "TELEFONO: "
                            End If
                        Else
                            sTelefono = "TELEFONO: "
                        End If
                        Set rs = Nothing
               '***************************************************************
               If Len(sTelefono) > 10 Then
                    sUserOld = rsCabecera!cUser
                    SaltaLinea sCadImp, nNumLineas
                    sCadImp = sCadImp & " Analista:" & rsCabecera!cUser & lnSaltoLinDoc
                    'SaltaLinea sCadImp, nNumLineas
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    'SaltaLinea sCadImp, nNumLineas
                    sCadImp = sCadImp & ImpreFormat("CREDITO/CLIENTE", 38) & ImpreFormat("DIRECCION", 44) & ImpreFormat("PREST", 10) & Space(9) & ImpreFormat("CUO APROB", 13) & ImpreFormat("SALDO.K", 13) & Space(8) & ImpreFormat("MONTO COBR", 13) & lnSaltoLinDoc
                    'SaltaLinea sCadImp, nNumLineas
                    sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
                    With rsCabecera
                      'SaltaLinea sCadImp, nNumLineas
                      sCadImp = sCadImp & ImpreFormat(!Titular, 40) & "DOMIC: " & ImpreFormat(!Direccion, 40) & ImpreFormat(Format(!KDesembolsado, "#0.00"), 10) & ImpreFormat(!NroCuotaApr, 13, 0) & Space(9) & ImpreFormat(Format(!SaldoK, "#0.00"), 13) & ImpreFormat(!MontoCobrar, 13) & lnSaltoLinDoc
                      'SaltaLinea sCadImp, nNumLineas
                      sCadImp = sCadImp & ImpreFormat(Mid(!cCtaCod, 1, 3), 3) & "-" & ImpreFormat(Mid(!cCtaCod, 4, 2), 2) & "-" & ImpreFormat(Mid(!cCtaCod, 6, 3), 3) & "-" & ImpreFormat(Mid(!cCtaCod, 9, 1), 1) & "-" & ImpreFormat(Mid(!cCtaCod, 10, 9), 10) & Space(5) & "TRABA: " & sDireccionTrabajo & "  " & sTelefono & "  " & Trim(sConvenio) & " CRED:" & !cRFA & lnSaltoLinDoc
                      
                    End With
               
             'SaltaLinea sCadImp, nNumLineas
             
             'cargando las garantias
                sCadImp = sCadImp & ImpreFormat("GARANTE", 38) & ImpreFormat("DIRECC", 40) & lnSaltoLinDoc
                Set oCredDoc = New DCredDoc
                Set rsGarant = oCredDoc.ObtenerGarantes(rsCabecera!cCtaCod)
                Do Until rsGarant.EOF
                    sCadImp = sCadImp & ImpreFormat(rsGarant!cpersnombre, 38) & ImpreFormat(rsGarant!cPersDireccDomicilio, 40) & lnSaltoLinDoc
                    rsGarant.MoveNext
                Loop
                Set rsGarant = Nothing
                
                sCadImp = sCadImp & Space(66) & ImpreFormat("CUOTA", 15) & ImpreFormat("FEC.VENC", 15) & ImpreFormat("S.CUOTA", 15) & ImpreFormat("ATR.CUOTA", 15) & lnSaltoLinDoc
                        Set oCredDoc = New DCredDoc
                        Set rsDetalle = oCredDoc.ListaMoraXAnalistaDetalle(rsCabecera!cCtaCod, psFecIni)
                        Set oCredDoc = Nothing
                '        SaltaLinea sCadImp, nNumLineas
                        'sCadImp = sCadImp & Space(66) & "*************** CUOTAS PENDIENTES*************" & Chr(10)
                        'sCadImp = sCadImp & Chr(10)
                        Do Until rsDetalle.EOF
                 '           SaltaLinea sCadImp, nNumLineas
                            sCadImp = sCadImp & Space(64) & ImpreFormat(rsDetalle!nCuota, 6, 0) & Space(10) & ImpreFormat(rsDetalle!dFecha, 10) & Space(9) & ImpreFormat(Format(rsDetalle!SaldoCuota, "#0.00"), 10) & ImpreFormat(rsDetalle!AtrasoCuota, 10, 0) & lnSaltoLinDoc
                            rsDetalle.MoveNext
                        Loop
                        Set rsDetalle = Nothing
                'SaltaLinea sCadImp, nNumLineas
               sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
               End If
            Else
               ' es el mismo analista
               '--**************** DATOS ADICIONALES****************************
                        Set oCredDoc = New DCredDoc
                        sDireccionTrabajo = oCredDoc.ObtenerDireccTrabajoTitular(rsCabecera!cCtaCod)
                        sConvenio = oCredDoc.ObtenerConvenio(rsCabecera!cCtaCod)
                        Set rs = oCredDoc.ObtenerTelefonosTitular(rsCabecera!cCtaCod)
                        Set oCredDoc = Nothing
                        'sDireccionTrabajo = "Dir.Trabaj: " & sDireccionTrabajo
                        If Not rs.EOF And Not rs.BOF Then
                            If Not IsNull(rs!cPersTelefono) Then
                               ' bTelefono = True
                                If Len(Trim(rs!cPersTelefono)) > 0 Then
                                    sTelefono = "TELEFONO: " & Trim(rs!cPersTelefono)
                                Else
                                    sTelefono = "TELFONO: "
                                End If
                            ElseIf Not IsNull(Trim(rs!cPersTelefono2)) Then
'                                bTelefono = True
                                If Len(rs!cPersTelefono2) > 0 Then
                                    sTelefono = "TELEFONO: " & Trim(rs!cPersTelefono2)
                                End If
                            Else
                                sTelefono = "TELEFONO: "
                            End If
                        Else
                            sTelefono = "TELEFONO: "
                        End If
                        Set rs = Nothing
               '***************************************************************
               If Len(sTelefono) > 10 Then
                   With rsCabecera
                   ' SaltaLinea sCadImp, nNumLineas
                    sCadImp = sCadImp & ImpreFormat(Trim(!Titular), 40) & "DOMIC: " & ImpreFormat(Trim(!Direccion), 40) & ImpreFormat(Format(!KDesembolsado, "#0.00"), 10) & ImpreFormat(!NroCuotaApr, 13, 0) & Space(9) & ImpreFormat(Format(!SaldoK, "#0.00"), 13) & ImpreFormat(!MontoCobrar, 13) & lnSaltoLinDoc
                    'SaltaLinea sCadImp, nNumLineas
                    sCadImp = sCadImp & ImpreFormat(Mid(!cCtaCod, 1, 3), 3) & "-" & ImpreFormat(Mid(!cCtaCod, 4, 2), 2) & "-" & ImpreFormat(Mid(!cCtaCod, 6, 3), 3) & "-" & ImpreFormat(Mid(!cCtaCod, 9, 1), 1) & "-" & ImpreFormat(Mid(!cCtaCod, 10, 9), 10) & Space(9) & "TRABA: " & sDireccionTrabajo & "  " & sTelefono & "  " & Trim(sConvenio) & " CRED: " & !cRFA & lnSaltoLinDoc
                    
                    '********************************
                    'cargando las garantias
                    sCadImp = sCadImp & ImpreFormat("GARANTE", 38) & ImpreFormat("DIRECC", 40) & lnSaltoLinDoc
                    Set oCredDoc = New DCredDoc
                    Set rsGarant = oCredDoc.ObtenerGarantes(!cCtaCod)
                    Do Until rsGarant.EOF
                        sCadImp = sCadImp & ImpreFormat(rsGarant!cpersnombre, 38) & ImpreFormat(rsGarant!cPersDireccDomicilio, 40) & lnSaltoLinDoc
                        rsGarant.MoveNext
                    Loop
                    Set rsGarant = Nothing
                   End With
                   'SaltaLinea sCadImp, nNumLineas
                   
                   sCadImp = sCadImp & Space(66) & ImpreFormat("CUOTA", 15) & ImpreFormat("FEC.VENC", 15) & ImpreFormat("S.CUOTA", 15) & ImpreFormat("ATR.CUOTA", 15) & lnSaltoLinDoc
                            Set oCredDoc = New DCredDoc
                            Set rsDetalle = oCredDoc.ListaMoraXAnalistaDetalle(rsCabecera!cCtaCod, psFecIni)
                            Set oCredDoc = Nothing
                            'sCadImp = sCadImp & Space(66) & "*************** CUOTAS PENDIENTES*************" & Chr(10)
                            'sCadImp = sCadImp & Chr(10)
                            Do Until rsDetalle.EOF
                             '   SaltaLinea sCadImp, nNumLineas
                                sCadImp = sCadImp & Space(64) & ImpreFormat(rsDetalle!nCuota, 6, 0) & Space(10) & ImpreFormat(rsDetalle!dFecha, 10) & Space(5) & ImpreFormat(Format(rsDetalle!SaldoCuota, "#0.00"), 10) & ImpreFormat(rsDetalle!AtrasoCuota, 10, 0) & lnSaltoLinDoc
                                rsDetalle.MoveNext
                            Loop
                            Set rsDetalle = Nothing
                   'SaltaLinea sCadImp, nNumLineas
                   sCadImp = sCadImp & String(160, "-") & lnSaltoLinDoc
               End If
             End If
        rsCabecera.MoveNext
     Loop
     Set rsCabecera = Nothing
     RaiseEvent FinalizaBarra
     ImprimeMoraXAnalitaNewTelefono = sCadImp
End Function


Public Function Repor_Imprime108409(ByVal psMoneda As String, ByVal pdFechaInicial As Date, ByVal pdFechaFin As Date, _
                                    ByVal psNomAge As String, ByVal psCodUser As String, ByVal psNomCmac As String, _
                                    ByVal pMatAgencias As Variant, ByVal pMatInstitucion As Variant, _
                                    ByVal pdFecSis As Date) As String
 'creditos cancelados por institucion
 Dim I As Integer
 Dim J As Integer
 Dim nPv As Integer
 Dim sCadImp As String
 Dim nTotal As Integer
 Dim rs As ADODB.Recordset
 Dim sAgencias As String
 'Dim sInstitucion As String
 'Dim sAgencias As String
 Dim odCredDoc As DCredDoc
 Dim sAgencia As String
 Dim sInstitucion As String
 'Dim sCadImp As String
 'armando las matrices
'  For i = 0 To UBound(pMatAgencias) - 1
'    If i = 0 Then
'        sAgencias = "'" & pMatAgencias(i) & "'"
'    Else
'        sAgencias = sAgencias & ",'" & pMatAgencias(i) & "'"
'    End If
'  Next i
'
'  For i = 0 To UBound(pMatInstitucion) - 1
'    If i = 0 Then
'        sInstitucion = "'" & pMatInstitucion(i) & "'"
'    Else
'        sInstitucion = sInstitucion & ",'" & pMatInstitucion(i) & "'"
'    End If
'  Next i
'  Set oDCredDoc = New DCredDoc
'  Set rs = oDCredDoc.ListaCreditosConvenioCancaledos(pdFechaIni, pdFechaFin, psMoneda, psCodAgencia, psCodInstitucion)
'  Set oDCredDoc = Nothing
'
'  If Not rs.RecordCount > 0 Then
'    nTotal = rs.RecordCount
'    RaiseEvent IniciaBarra(nTotal)
'  Else
'    Repor_Imprime108409 = ""
'    Exit Function
'  End If
  
  Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "INFORME DE CREDITOS CANCELADOS POR INSTITUCION", psNomCmac, , True)
  sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
  
  For I = 0 To UBound(pMatAgencias) - 1
      sCadImp = sCadImp & lnSaltoLinDoc
      Set odCredDoc = New DCredDoc
      sAgencia = odCredDoc.GetAgencia(pMatAgencias(I))
      Set odCredDoc = Nothing
      
      sCadImp = sCadImp & " AGENCIA: " & sAgencia & lnSaltoLinDoc
      sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
      
      nPv = 0
      
      For J = 0 To UBound(pMatInstitucion) - 1
         Set odCredDoc = New DCredDoc
         sInstitucion = odCredDoc.GetPersona(pMatInstitucion(J))
         Set odCredDoc = Nothing
         
         sCadImp = sCadImp & "INSTITUCION CON CONVENIO:" & sInstitucion & lnSaltoLinDoc
         sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
             
         'imprimiendo cabeceras
         sCadImp = sCadImp & ImpreFormat("NRO CREDITO", 20) & ImpreFormat("MODULAR", 10) & ImpreFormat("TITULAR", 30) & ImpreFormat("MONTO DESEMBOLSO", 18) & ImpreFormat("FECHA CANCELACION", 20) & lnSaltoLinDoc
         sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
         
         'Insertando el Detalle
          Set odCredDoc = New DCredDoc
          Set rs = odCredDoc.ListaCreditosConvenioCancaledos(pdFechaInicial, pdFechaFin, psMoneda, pMatAgencias(I), pMatInstitucion(J))
          Set odCredDoc = Nothing
          
          nTotal = rs.RecordCount
          
          If nTotal > 0 Then
             RaiseEvent IniciaBarra(nTotal)
             
             Do Until rs.EOF
                RaiseEvent ProgresoBarra(nPv, "LISTA DE CANCELADOS POR INSTITUCION", "PROCESANDO...")
                sCadImp = sCadImp & ImpreFormat(rs!cCtaCod, 20) & ImpreFormat(rs!cCodModular, 10) & ImpreFormat(rs!cTitular, 30) & ImpreFormat(rs!nMontoCol, 18) & ImpreFormat(rs!dFecha, 20) & lnSaltoLinDoc
                nPv = nPv + 1
                rs.MoveNext
             Loop
             Set rs = Nothing
          End If
          Set rs = Nothing
          sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
      Next J
      RaiseEvent FinalizaBarra
      sCadImp = sCadImp & String(130, "-") & lnSaltoLinDoc
  Next I
  Repor_Imprime108409 = sCadImp
End Function

Public Function Repor_Impre108410(ByVal pdFechaInicial As Date, ByVal pdFechaFin As Date, _
                                  ByVal psNomAge As String, ByVal psCodUser As String, ByVal psNomCmac As String, _
                                  ByVal pMatAgencias As Variant, ByVal pMatInstitucion As Variant, _
                                  ByVal pdFecSis As Date) As String
                                  
'Creditos aprobados por Institucion

    Dim I As Integer
    Dim J As Integer
    Dim nPv As Integer
    Dim nTotal As Integer
    
    Dim sCadImp As String
    Dim sAgencia As String
    Dim sInstitucion As String
    
    Dim rs As ADODB.Recordset
    Dim odCredDoc As DCredDoc
    
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "INFORME DE CREDITOS DESEMBOLSADOS POR INSTITUCION", psNomCmac, , True)
    sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
    
    
    For I = 0 To UBound(pMatAgencias) - 1
        sCadImp = sCadImp & lnSaltoLinDoc
        
        Set odCredDoc = New DCredDoc
        sAgencia = odCredDoc.GetAgencia(pMatAgencias(I))
        Set odCredDoc = Nothing
        
        sCadImp = sCadImp & "AGENCIA: " & sAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
        
        nPv = 0
        
        For J = 0 To UBound(pMatInstitucion) - 1
            Set odCredDoc = New DCredDoc
            sInstitucion = odCredDoc.GetPersona(pMatInstitucion(J))
            Set odCredDoc = Nothing
            
            sCadImp = sCadImp & "INSTITUCION CON CONVENIO:" & sInstitucion & lnSaltoLinDoc
            sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
            
            'Impriendo Cabeceras
            sCadImp = sCadImp & ImpreFormat("NRO CREDITO", 20) & ImpreFormat("MODULAR", 10) & ImpreFormat("TITUlAR", 30) & ImpreFormat("MONTO DESEMBOLSO", 18) & ImpreFormat("FECHA DESEMBOLSO", 20) & ImpreFormat("Nro Cuota", 10) & ImpreFormat("Mont. Cuota", 15) & ImpreFormat("Fec 1 Cuota", 15) & lnSaltoLinDoc
            sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
            
            'Insertanndo el Detalle
            Set odCredDoc = New DCredDoc
            
            Set rs = odCredDoc.ListaCreditosXInstitucionDesembolsados(pMatAgencias(I), pMatInstitucion(J), pdFechaInicial, pdFechaFin)
            Set odCredDoc = Nothing
            nTotal = rs.RecordCount
            If nTotal > 0 Then
                RaiseEvent IniciaBarra(rs.RecordCount)
                
                Do Until rs.EOF
                    RaiseEvent ProgresoBarra(nPv, "LISTA DE DESEMBOLSADOS POR INSTITUCION", "PROCESANDO...")
                    sCadImp = sCadImp & ImpreFormat(rs!cCtaCod, 20) & ImpreFormat(rs!cCodModular, 10) & ImpreFormat(rs!cpersnombre, 30) & ImpreFormat(rs!nMontoCol, 18) & ImpreFormat(rs!Fecha, 20) & ImpreFormat(rs!NumeroCuota, 4, 0) & ImpreFormat(rs!MontoCuota, 15, 2) & Space(6) & ImpreFormat(Format(rs!FechaVenc, "dd/mm/yyyy"), 15) & lnSaltoLinDoc
                    nPv = nPv + 1
                    rs.MoveNext
                Loop
                Set rs = Nothing
                sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
            End If
        Next J
        RaiseEvent FinalizaBarra
        sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
    Next I
    Repor_Impre108410 = sCadImp
End Function
Public Function Repo_108118_SaldosDiariosCred(ByVal pdFechaInicial As Date, _
                                  ByVal psNomAge As String, ByVal psCodUser As String, ByVal psNomCmac As String, _
                                  ByVal pdFecSis As Date, ByVal pdCodAge As String) As String
                                  
'Creditos aprobados por Institucion

    Dim I As Integer
    Dim J As Integer
    Dim nPv As Integer
    Dim nTotal As Integer
    
    Dim sCadImp As String
    Dim sAgencia As String
    Dim sInstitucion As String
    
    Dim rs As ADODB.Recordset
    Dim odCredDoc As DCredDoc
    
    
    Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE DE SALDOS DIARIOS DE CREDITOS AL:" & Format(pdFechaInicial, "DD/MM/YYYY"), psNomCmac, , True)
    sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
    
    
    'For i = 0 To UBound(pMatAgencias) - 1
        sCadImp = sCadImp & lnSaltoLinDoc
        
        Set odCredDoc = New DCredDoc
        sAgencia = odCredDoc.GetAgencia(pdCodAge)
        Set odCredDoc = Nothing
        
        sCadImp = sCadImp & "AGENCIA: " & sAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
        
        nPv = 0
        
        'For J = 0 To UBound(pMatInstitucion) - 1
            'Set odCredDoc = New DCredDoc
            'sInstitucion = odCredDoc.GetPersona(pMatInstitucion(J))
            'Set odCredDoc = Nothing
            
            'sCadImp = sCadImp & "INSTITUCION CON CONVENIO:" & sInstitucion & lnSaltoLinDoc
            sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
            
            'Impriendo Cabeceras
            sCadImp = sCadImp & ImpreFormat("Tipo Producto", 23) & ImpreFormat("Cred. Soles", 15) & ImpreFormat("Cred. Dolares", 15) & _
             ImpreFormat("Jud Vig Soles", 13) & ImpreFormat("Jud Vig Dolares", 13) & ImpreFormat("Jud Cast Soles", 13) & ImpreFormat("Jud Cast Dolares", 13) & _
             ImpreFormat("Total Soles", 13) & ImpreFormat("Total Dolares", 13) & lnSaltoLinDoc
            sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
            
            'Insertanndo el Detalle
            Set odCredDoc = New DCredDoc
            
            Set rs = odCredDoc.ColCredRepCredSaldosDiarios(pdCodAge, pdFechaInicial)
            Set odCredDoc = Nothing
            'nTotal = rs.RecordCount
            'If nTotal > 0 Then
                RaiseEvent IniciaBarra(rs.RecordCount)
                
                Do Until rs.EOF
                    RaiseEvent ProgresoBarra(nPv, "SALDOS DIARIOS DE COLOCACIONES", "PROCESANDO...")
'                JDolares              JCSoles               JCDolares             TSoles                TDolares
                    sCadImp = sCadImp & ImpreFormat(rs!nConsValor & " " & rs!cConsDescripcion, 20) & _
                    ImpreFormat(rs!CSoles, 13) & ImpreFormat(rs!CDolares, 13) & _
                    ImpreFormat(rs!JSoles, 13) & ImpreFormat(rs!JDolares, 13) & _
                    ImpreFormat(rs!JCSoles, 13) & ImpreFormat(rs!JCDolares, 13) & _
                    ImpreFormat(rs!TSoles, 13) & ImpreFormat(rs!TDolares, 13) & lnSaltoLinDoc
                    nPv = nPv + 1
                    rs.MoveNext
                Loop
                Set rs = Nothing
                sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
            'End If
        'Next J
        RaiseEvent FinalizaBarra
        sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
    'Next i
    Repo_108118_SaldosDiariosCred = sCadImp
End Function
Public Sub CuadroComisiones(ByVal pdFechaInicial As Date, ByVal pdFechaFinal As Date, ByVal pnTipoCambio As Double, _
ByVal psCodAgen As String, ByVal pNomAge As String, ByVal psCodUser As String)
    Dim rs As ADODB.Recordset
    Dim odCredDoc As DCredDoc
    Dim m_Excel As Excel.Application
    Dim oLibroExcel As Excel.Workbook
    Dim oHojaExcel As Excel.Worksheet
    Dim nFila As Integer
    
    Set m_Excel = New Excel.Application
    
    'seteando la aplicacion
    Set m_Excel = New Excel.Application
    Set oLibroExcel = m_Excel.Workbooks.Add
    Set oHojaExcel = oLibroExcel.Worksheets(1)
    oHojaExcel.Visible = xlSheetVisible
    
    oHojaExcel.Activate
    
    oHojaExcel.PageSetup.Zoom = 80
    oHojaExcel.PageSetup.Orientation = xlLandscape
    m_Excel.Range("A1:R100000").Font.Size = 7
    m_Excel.Selection.NumberFormat = "#,#,##0.00"
    
    'creando encabezado del reporte
    oHojaExcel.Range("A2:D2").Merge
    oHojaExcel.Range("A2:D2").value = "CUADRO DE COMISIONES"
    oHojaExcel.Range("A2:D2").Font.Italic = True
    oHojaExcel.Range("A2:D2").Font.Size = 13
    
    oHojaExcel.Range("A3:D2").Merge
    oHojaExcel.Range("A3:D2").value = "CAJA MUNICIPAL DE ICA S.A"
    oHojaExcel.Range("A3:D2").Font.Italic = True
    oHojaExcel.Range("A3:D2").Font.Size = 13
    
    oHojaExcel.Range("A4:D2").Merge
    oHojaExcel.Range("A4:D2").value = "Nombre de la Agencia:" & pNomAge
    oHojaExcel.Range("A4:D2").Font.Italic = True
    oHojaExcel.Range("A4:D2").Font.Size = 13
    
    oHojaExcel.Range("A5:D2").Merge
    oHojaExcel.Range("A5:D2").value = "Usuario:" & psCodUser
    oHojaExcel.Range("A5:D2").Font.Italic = True
    oHojaExcel.Range("A5:D2").Font.Size = 13

    'armando  las columnas de cuadro de Creditos nuevos
    With oHojaExcel
        .Cells(7, 3).Font.Bold = True
        .Cells(7, 3).Font.Italic = True
        .Cells(7, 3) = "Cuadro de Creditos Nuevos"
    End With
    
    nFila = 8
    
    'Cuadro de Creditos nuevos
    
End Sub

Public Function ListaCredNoDesembolsados(ByVal psCodAgen As String, ByVal pdFechaIni As Date, _
ByVal pdFechaFin As Date, ByVal psNomAge As String, ByVal pdFechaSis As Date, _
ByVal psCodUser As String, ByVal psNomCmac As String) As String
    
    Dim rs As ADODB.Recordset
    Dim nTotal As Integer
    Dim psCadImp As String
    Dim odCredDoc As DCredDoc
    
        Set odCredDoc = New DCredDoc
        Set rs = odCredDoc.ListaCredNoDesembolsados(psCodAgen, pdFechaIni, pdFechaFin)
        Set odCredDoc = Nothing
        
        If rs.RecordCount > 0 Then
            nTotal = rs.RecordCount
        Else
            Exit Function
        End If
        
        RaiseEvent IniciaBarra(nTotal)
        Call ImprimeCabeceraDocumento(psCadImp, psNomAge, pdFechaSis, psCodUser, "Lista de Creditos No Desembolsado", psNomCmac)
        
        'armando las columnas
        psCadImp = psCadImp & ImpreFormat("Credito", 20) & ImpreFormat("Nombre ", 25) & ImpreFormat("Monto", 8, 2) & ImpreFormat("Direccion", 20) & ImpreFormat("Telefono", 10) & ImpreFormat("Moneda", 6) & ImpreFormat("Usuario", 6) & lnSaltoLinDoc
        
        Do Until rs.EOF
        psCadImp = psCadImp & ImpreFormat(rs!cCtaCod, 20) & ImpreFormat(rs!cpersnombre, 25) & ImpreFormat(rs!nMonto, 8, 2) & ImpreFormat(rs!cDireccion, 20) & ImpreFormat(rs!cTelefono, 10) & ImpreFormat(rs!Moneda, 6) & ImpreFormat(rs!cUser, 6) & lnSaltoLinDoc
            rs.MoveNext
        Loop
        Set rs = Nothing
        ListaCredNoDesembolsados = psCadImp
End Function
    

Function ObtenerProxCuota(ByVal psCtaCod As String) As Recordset
    Dim sSQL As String
    Dim oConec As DConecta
    
        
    Set oConec = New DConecta
    oConec.AbreConexion
    
    sSQL = "Select CC.nCuota,CC.dVenc"
    sSQL = sSQL & " From ColocCalendario CC"
    sSQL = sSQL & "       Inner Join ColocacCred C on CC.cCtaCod=C.cCtaCod and CC.nNroCalen=C.nNroCalen"
    sSQL = sSQL & " Where CC.cCtaCod='" & psCtaCod & "' and CC.nColocCalendApl=1 and CC.nColocCalendEstado=0 and "
    sSQL = sSQL & "       CC.nCuota=(Select Min(CC.nCuota)"
    sSQL = sSQL & "             From ColocCalendario CC"
    sSQL = sSQL & "               Inner Join ColocacCred C on CC.cCtaCod=C.cCtaCod and CC.nNroCalen=C.nNroCalen"
    sSQL = sSQL & "         Where CC.cCtaCod='" & psCtaCod & "' and CC.nColocCalendApl=1 and CC.nColocCalendEstado=0)             "
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set ObtenerProxCuota = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Function ObtenerCuotaPagada(ByVal psCtaCod As String) As String
    Dim sSQL As String
    Dim oConec As DConecta
    Dim rs As ADODB.Recordset
    
'    sSql = "Select Max(nCuota) as nCuota"
'    sSql = sSql & " From ColocCalendario C"
'    sSql = sSql & " Inner Join ColocacCred CC on C.cCtaCod=CC.cCtaCod and C.nNroCalen=CC.nNroCalen"
'    sSql = sSql & " Where C.cCtaCod='" & psCtaCod & "' and C.nColocCalendApl=1 and C.nColocCalendEstado=1"
    
    sSQL = "Select Max(MCD.nNroCuota) as nCuota"
    sSQL = sSQL & " From MovColDet MCD"
    sSQL = sSQL & " Where MCD.nMovNro=(Select Max(M.nMovNro)"
    sSQL = sSQL & "           From MovCol MC"
    sSQL = sSQL & "                Inner Join Mov M on M.nMovNro=MC.nMovNro"
    sSQL = sSQL & "           Where MC.cCtaCod='" & psCtaCod & "' and M.nMovFlag=0) and"
    sSQL = sSQL & " MCD.cOpeCod not Like '107[123456789]%'" & " and MCD.cCtaCod='" & psCtaCod & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    
    If Not rs.EOF And Not rs.BOF Then
       ObtenerCuotaPagada = IIf(IsNull(rs!nCuota), 1, rs!nCuota)
    End If
    Set rs = Nothing
End Function


Function EstadisRFCDIF(ByVal pdFechaInicio As Date, pdFechaIniciPag As Date, pdFechaFinPag As Date) As String
    Dim rs As ADODB.Recordset
    Dim odCredDoc As DCredDoc
    
    Dim m_Excel As Excel.Application
    Dim oLibroExcel As Excel.Workbook
    Dim oHojaExcel As Excel.Worksheet
    
    Dim I As Integer
    Dim nFila As Integer
    
    
    
    Set odCredDoc = New DCredDoc
    Set rs = odCredDoc.Repo_108506(pdFechaInicio, pdFechaIniciPag, pdFechaFinPag)
    Set odCredDoc = Nothing
    
    
    If rs.RecordCount > 0 Then
        RaiseEvent IniciaBarra(rs.RecordCount)
    Else
        Set rs = Nothing
        EstadisRFCDIF = ""
        Exit Function
    End If
    
    
    Set m_Excel = New Excel.Application
    m_Excel.Visible = True
    
    Set oLibroExcel = m_Excel.Workbooks.Add
    Set oHojaExcel = oLibroExcel.Worksheets(1)
    oHojaExcel.Visible = xlSheetVisible
    
    oHojaExcel.Activate
    
    oHojaExcel.PageSetup.Zoom = 80
    oHojaExcel.PageSetup.Orientation = xlLandscape
    m_Excel.Range("A1:R1000").Font.Size = 9
    
    
    m_Excel.Selection.NumberFormat = "#,##0.00"
    
    'Creando encabezado de inforne
    oHojaExcel.Range("A2:D2").Merge
   oHojaExcel.Range("A2.D2").value = "ESTADISTICA DE RFC/DIF " & MonthName(Month(pdFechaIniciPag))
   oHojaExcel.Range("A2:D2").Font.Italic = True
   oHojaExcel.Range("A2:D2").Font.Size = 13
   
   nFila = 3
   
   nFila = nFila + 2
   
   
   With oHojaExcel
    .Cells(nFila, 1) = MonthName(Month(pdFechaIniciPag))
    .Cells(nFila, 1).Font.Bold = True
    
    .Cells(nFila, 2) = "Producto"
    .Cells(nFila, 2).Font.Bold = True
    
    .Cells(nFila, 3) = "Agencia"
    .Cells(nFila, 3).Font.Bold = True
    
    .Cells(nFila, 4) = "Concepto"
    .Cells(nFila, 4).Font.Bold = True
    
    .Cells(nFila, 5) = "Saldo Inicial"
    .Cells(nFila, 5).Font.Bold = True
    
    .Cells(nFila, 6) = "Pago Mes"
    .Cells(nFila, 6).Font.Bold = True
    
    .Cells(nFila, 7) = "Saldo Final"
    .Cells(nFila, 7).Font.Bold = True
    
    .Range("D1").EntireColumn.NumberFormat = "0.00"
    .Range("E1").EntireColumn.NumberFormat = "0.00"
    .Range("F1").EntireColumn.NumberFormat = "0.00"
    .Range("G1").EntireColumn.NumberFormat = "0.00"
   End With
   
   
   
   
   nFila = nFila + 1
   
   rs.MoveFirst
   
   Do Until rs.EOF
      With oHojaExcel
            .Cells(nFila, 2) = rs!cConsDescripcion
            .Cells(nFila, 3) = rs!cAgeDescripcion
            .Cells(nFila, 4) = rs!cRFA
            .Cells(nFila, 5) = Format(rs!nMontoIni, "#0.00")
            .Cells(nFila, 6) = Format(rs!nMontoFin, "#0.00")
            .Cells(nFila, 7) = Format(rs!nDif, "#0.00")
      End With
      RaiseEvent ProgresoBarra(I, "Estadistica de RFC/DIF", "Cargando...")
      nFila = nFila + 1
    rs.MoveNext
   Loop
   Set rs = Nothing
   RaiseEvent FinalizaBarra
   Set m_Excel = Nothing
   EstadisRFCDIF = "OK"
End Function


Public Sub Repo_108314(ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, ByVal pnTipoCambio As Double, _
ByVal psCodAgen As String, Optional ByVal psCodUser As String)

    Dim rs As ADODB.Recordset
    Dim nMontoFiltro As Double
    Dim sSQL As String
    Dim oConec As DConecta
    Dim odCredDoc As DCredDoc
    
    Dim m_Excel As Excel.Application
    Dim oLibroExcel As Excel.Workbook
    Dim oHojaExcel As Excel.Worksheet
    
    Dim nFila As Integer
    
    Dim sAgencia As String
    
    sSQL = "Select nParamValor"
    sSQL = sSQL & " From ColocParametro"
    sSQL = sSQL & " Where nParamVar=5001"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        nMontoFiltro = rs!nParamValor
    End If
    Set rs = Nothing
    
    sSQL = "Select cAgeDescripcion"
    sSQL = sSQL & "From Agencias"
    sSQL = sSQL & " Where cAgeCod='" & psCodAgen & "'"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set rs = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
    
    If Not rs.EOF And Not rs.BOF Then
        sAgencia = rs!cAgeDescripcion
    End If
    Set rs = Nothing
    
    RaiseEvent IniciaBarra(30)
    
    Call EncabezadoExcel(m_Excel, oLibroExcel, oHojaExcel, "Informe de Creditos para Comisiones Mes -" & MonthName(Month(pdFechaFin)))
    
    'Creacion del Encabezado
    nFila = 3
    
    With oHojaExcel
        .Cells(nFila, 1) = "Usuario: " & psCodUser
        nFila = nFila + 1
        .Cells(nFila, 1) = "Filtro fecha : Desde " & pdFechaIni & " Hasta " & pdFechaFin
        nFila = nFila + 1
        .Cells(nFila, 1) = "Agencia: " & sAgencia
    End With
    nFila = nFila + 2
    
    
End Sub


Sub EncabezadoExcel(ByRef pm_Excel As Excel.Application, ByRef pLibroExcel As Excel.Workbook, _
ByRef pHojaExcel As Excel.Worksheet, ByRef psMensaje As String)

 Set pm_Excel = New Excel.Application

 Set pLibroExcel = pm_Excel.Workbooks.Add
 Set pHojaExcel = pLibroExcel.Worksheets(1)
 pHojaExcel.Visible = xlSheetVisible
 
 pHojaExcel.Activate
 
    pHojaExcel.PageSetup.Zoom = 80
    pHojaExcel.PageSetup.Orientation = xlLandscape
    pm_Excel.Range("A1:R1000").Font.Size = 8
    
    pm_Excel.Selection.NumberFormat = "#,###0.00"
 
    pHojaExcel.Range("A2:D2").Merge
    pHojaExcel.Range("A2:D2").value = psMensaje
    pHojaExcel.Range("A2:D2").Font.Italic = True
    pHojaExcel.Range("A2:D2").Font.Size = 13
End Sub


'ARCV 27-07-2006
Public Function ImprimeProyeccionCancelacionCreditos(ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
                                ByVal psCodUser As String, ByVal pdFecSis As Date, _
                                ByVal psNomAge As String, ByVal pMatProd As Variant, ByVal pMatAnalistas As Variant, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim nNumLineas As Integer

Dim scCtaCodAnt As String
Dim snDatos As String
Dim scTipoAnt As String

Dim oDCreDocT As DCredDoc
Dim rs As ADODB.Recordset
Dim bTelefono As Boolean
Dim sDireccionTrabajo As String
Dim sConvenio As String
Dim nTotal As Integer
scCtaCodAnt = ""
    
    
    Set oDCred = New DCredDoc
    Set R = oDCred.ListaCreditosCancelacionProyectados(pdFecIni, pdFecFin, pMatAnalistas, pMatProd)
    Set oDCred = Nothing
    
    sCadImp = ""
    nNumLineas = 0
    
        
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "PROYECCION DE CANCELACION DE CREDITOS DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        
        sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!Analista) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!Producto) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(170, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("TITULAR", 34) & ImpreFormat("DIRECCION", 25) & ImpreFormat("TELEF.", 10, 0) & ImpreFormat("CUOTAS", 10) & ImpreFormat("MONTO APROB.", 10) & ImpreFormat("SALDO CAP.", 10) & ImpreFormat("ATRA.ACUM.", 8) & ImpreFormat("DESEMBOLSO", 10) & ImpreFormat("CANCELACION", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(170, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
        'FINAL DE TITULAR
        If Trim(R!Cuenta) = "ZZ" Then
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            'FINAL DE Producto
            If R!Producto = "ZZ" Then
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL DEL REPORTE
                If R!Analista = "ZZ" Then
                    sCadImp = sCadImp & lsTab & String(160, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    Exit Do
                    'R.MoveNext
                    'If R.EOF Then
                    '    Exit Do
                    'End If
                Else
                    sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!Analista) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!Producto) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(170, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("TITULAR", 34) & ImpreFormat("DIRECCION", 25) & ImpreFormat("TELEF.", 10, 0) & ImpreFormat("CUOTAS", 10) & ImpreFormat("MONTO APROB.", 10) & ImpreFormat("SALDO CAP.", 10) & ImpreFormat("ATRA.ACUM.", 8) & ImpreFormat("DESEMBOLSO", 10) & ImpreFormat("CANCELACION", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(170, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "PRODUCTO : " & Trim(R!Producto) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(170, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("TITULAR", 34) & ImpreFormat("DIRECCION", 25) & ImpreFormat("TELEF.", 10, 0) & ImpreFormat("CUOTAS", 10) & ImpreFormat("MONTO APROB.", 10) & ImpreFormat("SALDO CAP.", 10) & ImpreFormat("ATRA.ACUM.", 8) & ImpreFormat("DESEMBOLSO", 10) & ImpreFormat("CANCELACION", 12) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(170, "-") & lnSaltoLinDoc
            End If
        Else
            
        End If
        
        'sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("TITULAR", 34) & ImpreFormat("DIRECCION", 25) & ImpreFormat("TELEF.", 10) & ImpreFormat("PLAZO", 10) & ImpreFormat("MONTO APROB.", 12) & ImpreFormat("SALDO CAP.", 12) & ImpreFormat("ATRA. ACUM.", 8) & ImpreFormat("DESEMBOLSO", 12) & ImpreFormat("CANCELACION", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat(R!Cuenta, 18) & ImpreFormat(R!Titular, 30) & ImpreFormat(R!Direccion, 30) & ImpreFormat(IIf(IsNull(R!Telefono), "", R!Telefono), 12) & ImpreFormat(R!NroCuotas, 5, 0) & ImpreFormat(R!Moneda, 4) & ImpreFormat(R!MontoAprob, 8) & ImpreFormat(R!Saldo, 8) & ImpreFormat(R!DiasAtrasoAcum, 8, 0) & ImpreFormat(R!Desembolso, 12) & ImpreFormat(R!Vencimiento, 12) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    'RaiseEvent FinalizaBarra
    
    ImprimeProyeccionCancelacionCreditos = sCadImp
  
End Function

'ARCV 28-10-2006
Public Function ImprimeListadoVigentesXInstitucion(ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatInstituciones As Variant, ByVal pMatAgencias As Variant, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosVigentesXInstituciones(pdFecSis, pMatInstituciones, pMatAgencias)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "LISTADO DE CREDITOS VIGENTES POR INSTITUCION AL: " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 35) & ImpreFormat("MONTO CREDITO", 12) & ImpreFormat("NRO MESES", 10) & ImpreFormat("CUOTA", 6) & ImpreFormat("PROX VCTO", 10) & ImpreFormat("MONTO", 10) & ImpreFormat("SALDO CAP.", 10) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE INSTITUCION
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL INSTITUCION ", 26)
            sCadImp = sCadImp & ImpreFormat("CASOS : " & R!nCasos, 30) & ImpreFormat(Format(R!nMontoDesemb, "#0.00"), 45)
            sCadImp = sCadImp & ImpreFormat(Format(R!nSaldoCuota, "#0.00"), 10) & ImpreFormat(Format(R!nSaldoCap, "#0.00"), 10) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE AGENCIA
            If R!cInstitucion = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("*** SUB TOTAL AGENCIA ", 26)
                sCadImp = sCadImp & ImpreFormat("CASOS : " & R!nCasos, 30) & ImpreFormat(Format(R!nMontoDesemb, "#0.00"), 45)
                sCadImp = sCadImp & ImpreFormat(Format(R!nSaldoCuota, "#0.00"), 10) & ImpreFormat(Format(R!nSaldoCap, "#0.00"), 10) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'TOTAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("*** TOTAL ", 26)
                    sCadImp = sCadImp & ImpreFormat("CASOS : " & R!nCasos, 30) & ImpreFormat(Format(R!nMontoDesemb, "#0.00"), 45)
                    sCadImp = sCadImp & ImpreFormat(Format(R!nSaldoCuota, "#0.00"), 10) & ImpreFormat(Format(R!nSaldoCap, "#0.00"), 10) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    Exit Do
                Else
                    nPuntPag = 0
                    sCadImp = sCadImp & Chr$(12)
                    
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 35) & ImpreFormat("MONTO CREDITO", 12) & ImpreFormat("NRO MESES", 10) & ImpreFormat("CUOTA", 6) & ImpreFormat("PROX VCTO", 10) & ImpreFormat("MONTO", 10) & ImpreFormat("SALDO CAP.", 10) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                
                End If
        
            Else
                sCadImp = sCadImp & lsTab & "INSTITUCION : " & Trim(R!cInstitucion) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("CLIENTE", 35) & ImpreFormat("MONTO CREDITO", 12) & ImpreFormat("NRO MESES", 10) & ImpreFormat("CUOTA", 6) & ImpreFormat("PROX VCTO", 10) & ImpreFormat("MONTO", 10) & ImpreFormat("SALDO CAP.", 10) & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!sTitular, 35) & _
                                    ImpreFormat(Format(R!nMontoDesemb, "#0.00"), 14) & ImpreFormat(R!nNroCuotas, 8) & ImpreFormat(R!nCuota - R!nSaldoMora, 6) & ImpreFormat(Format(R!dFecVencimiento, "dd/mm/yyyy"), 12) _
                                    & ImpreFormat(Format(R!nSaldoCuota, "#0.00"), 10) & ImpreFormat(Format(R!nSaldoCap, "#0.00"), 10) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeListadoVigentesXInstitucion = sCadImp

End Function

'ARCV 03-02-2007
Public Function ImprimeListadoCalificacion(ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatAgencias As Variant, ByVal pMatAnalistas As Variant, _
    ByVal pMatCalificacion As Variant, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaCreditosVigentesXCalificacion(pdFecSis, pMatAnalistas, pMatAgencias, pMatCalificacion)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE PARA CLASIFICACION AL: " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadCab = sCadImp
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!sAnalista) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(164, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CUENTA", 18) & ImpreFormat("NOMBRE", 30) & ImpreFormat("DIRECCION", 30) & ImpreFormat("M.OTORGADO", 10) & ImpreFormat("SALDO CAP.", 10) & ImpreFormat("CALIF.", 3) & ImpreFormat("ATR.PROM.", 9) & ImpreFormat("# CUO", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("MONT.CUOT.", 10) & ImpreFormat("FECHA VENCIM.", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(164, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE ANALISTA
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(164, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE AGENCIA
            If R!sAnalista = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(164, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'TOTAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & lnSaltoLinDoc
                    Exit Do
                Else
                    nPuntPag = 0
                    sCadImp = sCadImp & Chr$(12)
                    
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!sAnalista) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(164, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CUENTA", 18) & ImpreFormat("NOMBRE", 30) & ImpreFormat("DIRECCION", 30) & ImpreFormat("M.OTORGADO", 10) & ImpreFormat("SALDO CAP.", 10) & ImpreFormat("CALIF.", 3) & ImpreFormat("ATR.PROM.", 9) & ImpreFormat("# CUO", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("MONT.CUOT.", 10) & ImpreFormat("FECHA VENCIM.", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(164, "-") & lnSaltoLinDoc
                
                End If
        
            Else
                    sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!sAnalista) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(164, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CUENTA", 18) & ImpreFormat("NOMBRE", 30) & ImpreFormat("DIRECCION", 30) & ImpreFormat("M.OTORGADO", 10) & ImpreFormat("SALDO CAP.", 10) & ImpreFormat("CALIF.", 3) & ImpreFormat("ATR.PROM.", 9) & ImpreFormat("# CUO", 5) & ImpreFormat("PLAZO", 5) & ImpreFormat("MONT.CUOT.", 10) & ImpreFormat("FECHA VENCIM.", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(164, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 18) & ImpreFormat(R!sTitular, 30) & _
                                    ImpreFormat(R!sDireccion, 30) & ImpreFormat(R!nMontoDesemb, 8, 2) & ImpreFormat(R!nSaldoCap, 8, 2) & Space(2) & ImpreFormat(R!cClasificacion, 2) & ImpreFormat(R!nAtrasoProm, 4, 2) & Space(6) & ImpreFormat(R!sCuotaPend, 6) & ImpreFormat(Format(R!nPlazo, "#0"), 2) & ImpreFormat(R!nSaldoCuota, 8, 2) & Space(2) & ImpreFormat(Format(R!dFecVencimiento, "dd/mm/yyyy"), 10) _
                                     & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeListadoCalificacion = sCadImp

End Function

'05-02-2007
Public Function ImprimeListadoPolizas(ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatAgencias As Variant, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaDatosListadoPoliza(pdFecSis, pMatAgencias)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CONTROL DE POLIZAS AL: " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadCab = sCadImp
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "CLIENTE : " & Trim(R!sCliente) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("CUENTA", 18) & ImpreFormat("M.DESEMB.", 10) & ImpreFormat("SALDO CAP.", 9) & ImpreFormat("MON.", 2) & ImpreFormat("F.DESEMB.", 10) & ImpreFormat("F.TERMINO", 10) & ImpreFormat("POLIZA", 8) & ImpreFormat("CIA.SEGUROS", 12) & ImpreFormat("F.VIGENC.", 10) & ImpreFormat("F.VENCIM", 10) & ImpreFormat("CONDIC.", 8) & ImpreFormat("DIRECC. INMUEBLE", 25) & ImpreFormat("SUMA ASEG.", 10) & lnSaltoLinDoc
        sCadImp = sCadImp & String(172, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE ANALISTA
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE AGENCIA
            If R!sCliente = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'TOTAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & lnSaltoLinDoc
                    Exit Do
                Else
                    nPuntPag = 0
                    sCadImp = sCadImp & Chr$(12)
                    
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "CLIENTE : " & Trim(R!sCliente) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("CUENTA", 20) & ImpreFormat("M.DESEMB.", 10) & ImpreFormat("SALDO CAP.", 10) & ImpreFormat("MONEDA", 3) & ImpreFormat("F.DESEMB.", 10) & ImpreFormat("F.TERMINO", 10) & ImpreFormat("POLIZA", 5) & ImpreFormat("CIA.SEGUROS", 10) & ImpreFormat("F.VIGENC.", 12) & ImpreFormat("F.VENCIM", 12) & ImpreFormat("CONDICION", 8) & ImpreFormat("DIRECC. INMUEBLE", 25) & ImpreFormat("SUMA ASEG.($)", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
                End If
        
            Else
                    sCadImp = sCadImp & lsTab & "CLIENTE : " & Trim(R!sCliente) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("CUENTA", 20) & ImpreFormat("M.DESEMB.", 10) & ImpreFormat("SALDO CAP.", 10) & ImpreFormat("MONEDA", 3) & ImpreFormat("F.DESEMB.", 10) & ImpreFormat("F.TERMINO", 10) & ImpreFormat("POLIZA", 5) & ImpreFormat("CIA.SEGUROS", 10) & ImpreFormat("F.VIGENC.", 12) & ImpreFormat("F.VENCIM", 12) & ImpreFormat("CONDICION", 8) & ImpreFormat("DIRECC. INMUEBLE", 25) & ImpreFormat("SUMA ASEG.($)", 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
            End If
        End If
        
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 18) & ImpreFormat(R!nMontoDesemb, 8, 2) & ImpreFormat(R!nSaldoCap, 8, 2) & Space(1) & ImpreFormat(R!cMoneda, 2) & ImpreFormat(Format(R!dDesembolso, "dd/mm/yyyy"), 10) & ImpreFormat(Format(R!dTermino, "dd/mm/yyyy"), 10) & ImpreFormat(R!cCodPolizaAseg, 8) & ImpreFormat(R!cCiaSeguros, 12) & ImpreFormat(Format(R!dVigenciaAseg, "dd/mm/yyyy"), 10) & ImpreFormat(Format(R!dVencAseg, "dd/mm/yyyy"), 10) & ImpreFormat(R!cCondicionPoliza, 8) & ImpreFormat(R!cDireccionInmueble, 25) & ImpreFormat(R!nSumaAsegurada, 8, 2) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeListadoPolizas = sCadImp

End Function

'07-02-2007
Public Function ImprimeListadoGruposEconomicos(ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaDatosGruposEconomicos()
    Set oDCred = Nothing

    If R.RecordCount > 0 Then
        R.MoveFirst
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "LISTADO DE GRUPOS ECONOMICOS AL: " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadCab = sCadImp
        sCadImp = sCadImp & lsTab & "GRUPO : " & Trim(R!cGrupoEcon) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("EMPRESAS VINCULADAS", 30) & ImpreFormat("TIPO  DOC.  NUMERO  ", 21) & ImpreFormat("DOMICILIO", 25) & ImpreFormat("VINC.", 10) & ImpreFormat("ACCIONISTAS,DIRECTORES,ETC", 26) & ImpreFormat("NRO DOC.", 11) & ImpreFormat("RESIDENCIA", 20) & ImpreFormat("ACCION.", 8) & ImpreFormat("CARGO", 12) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
        R.MoveNext
    End If

    Do While Not R.EOF

        If Trim(R!nOrden) = 1 Then
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "GRUPO : " & Trim(R!cGrupoEcon) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("EMPRESAS VINCULADAS", 30) & ImpreFormat("TIPO  DOC.  NUMERO  ", 21) & ImpreFormat("DOMICILIO", 25) & ImpreFormat("VINC.", 10) & ImpreFormat("ACCIONISTAS,DIRECTORES,ETC", 26) & ImpreFormat("NRO DOC.", 11) & ImpreFormat("RESIDENCIA", 20) & ImpreFormat("ACCION.", 8) & ImpreFormat("CARGO", 12) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
        Else
            sCadImp = sCadImp & ImpreFormat(R!cPersNombreCli, 30) & ImpreFormat(R!cTipoPer & "   " & R!cTipoDoc & "  " & R!cNumeroDoc, 21) & ImpreFormat(R!cDomicilio, 25) & ImpreFormat(R!cVinculacion, 10) & ImpreFormat(R!cPersNombreOtro, 26) & ImpreFormat(R!cNumeroDocOtro, 11) & ImpreFormat(R!cDomicilioOtro, 20) & ImpreFormat(R!nPorcenOtro, 4, 2) & ImpreFormat(R!cCargoOtro, 12) & lnSaltoLinDoc
        End If

        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeListadoGruposEconomicos = sCadImp
End Function

Public Function ImprimeDetalleGarantiasEnSoles(ByVal psCodUser As String, ByVal pdFecSis As Date, _
            ByVal psNomAge As String, ByVal pMatAgencias As Variant, ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, _
            Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaDetalleGarantiasEnSoles(pMatAgencias, pdFechaIni, pdFechaFin)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE DE GARANTIAS (En Soles) AL: " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadCab = sCadImp
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("FECHA", 12) & ImpreFormat("RESP.", 5) & ImpreFormat("NOMBRE", 30) & ImpreFormat("OTORGADO", 11) & ImpreFormat("PREFERIDAS", 11) & ImpreFormat("NO PREFERID.", 11) & ImpreFormat("PLAZO FIJO", 11) & Space(4) & ImpreFormat("TOTAL", 11) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE ANALISTA
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("TOTAL AGENCIA", 15) & ImpreFormat("", 12) & ImpreFormat("", 5) & ImpreFormat("", 30) & ImpreFormat(R!nMontoDesemb, 11, 2) & ImpreFormat(R!nPreferidas, 11, 2) & ImpreFormat(R!nNoPreferidas, 11, 2) & ImpreFormat(R!nAutoliquidables, 11, 2) & ImpreFormat(R!nTotal, 11, 2) & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE AGENCIA
            If R!sAgencia = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("TOTAL REPORTE", 15) & ImpreFormat("", 12) & ImpreFormat("", 5) & ImpreFormat("", 30) & ImpreFormat(R!nMontoDesemb, 11, 2) & ImpreFormat(R!nPreferidas, 11, 2) & ImpreFormat(R!nNoPreferidas, 11, 2) & ImpreFormat(R!nAutoliquidables, 11, 2) & ImpreFormat(R!nTotal, 11, 2) & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'TOTAL FINAL
                If R!sAgencia = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(172, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("", 20) & ImpreFormat("", 12) & ImpreFormat("", 5) & ImpreFormat("", 30) & ImpreFormat(R!nMontoDesemb, 11, 2) & ImpreFormat(R!nPreferidas, 11, 2) & ImpreFormat(R!nNoPreferidas, 11, 2) & ImpreFormat(R!nAutoliquidables, 11, 2) & ImpreFormat(R!nTotal, 11, 2) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    Exit Do
                Else
                    nPuntPag = 0
                    sCadImp = sCadImp & Chr$(12)
                    
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("FECHA", 12) & ImpreFormat("RESP.", 5) & ImpreFormat("NOMBRE", 30) & ImpreFormat("OTORGADO", 11) & ImpreFormat("PREFERIDAS", 11) & ImpreFormat("NO PREFERID.", 11) & ImpreFormat("PLAZO FIJO", 11) & Space(4) & ImpreFormat("TOTAL", 11) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                End If
        
            Else
                    sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("CREDITO", 20) & ImpreFormat("FECHA", 12) & ImpreFormat("RESP.", 5) & ImpreFormat("NOMBRE", 25) & ImpreFormat("OTORGADO", 11) & ImpreFormat("PREFERIDAS", 11) & ImpreFormat("NO PREFERID.", 11) & ImpreFormat("PLAZO FIJO", 11) & ImpreFormat("TOTAL", 11) & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & String(145, "-") & lnSaltoLinDoc
            End If
        End If
        
        
        sCadImp = sCadImp & lsTab & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(Format(R!dFecha, "dd-mm-yyyy"), 12) & ImpreFormat(R!sAnalista, 5) & ImpreFormat(R!sTitular, 25) & ImpreFormat(R!nMontoDesemb, 11, 2) & ImpreFormat(R!nPreferidas, 11, 2) & ImpreFormat(R!nNoPreferidas, 11, 2) & ImpreFormat(R!nAutoliquidables, 11, 2) & ImpreFormat(R!nTotal, 11, 2) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeDetalleGarantiasEnSoles = sCadImp

End Function

Public Function ImprimeGarantiasXClientes(ByVal psCodUser As String, ByVal pdFecSis As Date, _
            ByVal psNomAge As String, ByVal pMatAgencias As Variant, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim sCadTemp  As String
Dim I As Integer
Dim sAgencia As String
Dim nTotalSoles As Double
Dim nTotalDolares As Double
Dim P As Integer, nLi As Integer

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaGarantiasXCliente(pMatAgencias)
    Set oDCred = Nothing
    If R.RecordCount > 0 Then
        R.MoveFirst
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE DE GARANTIAS POR CLIENTE AL: " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(110, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("CUENTA", 20) & ImpreFormat("ESTADO", 15) & ImpreFormat("MONEDA", 8) & ImpreFormat("MONTO DESEMB.", 12) & ImpreFormat("FECHA DESEMB.", 12) & ImpreFormat("ANALISTA", 8) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(110, "-") & lnSaltoLinDoc
        sCadCab = sCadImp
        sCadImp = sCadImp & lsTab & "CLIENTE : " & Trim(R!cTitular) & Space(4) & Trim(R!cpersnombre) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(110, "-") & lnSaltoLinDoc
        R.MoveNext
    End If
    P = 0
    nTotalSoles = 0
    nTotalDolares = 0

    Do While Not R.EOF
        P = P + 1
        If Trim(R!nOrden) = 1 Then
            sCadTemp = sCadTemp & lnSaltoLinDoc
            sCadTemp = sCadTemp & lsTab & String(110, "-") & lnSaltoLinDoc
            sCadTemp = sCadTemp & lsTab & "CLIENTE : " & Trim(R!cTitular) & Space(4) & Trim(R!cpersnombre) & lnSaltoLinDoc
        ElseIf Trim(R!nOrden) = 2 Then
            sCadTemp = sCadTemp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(R!cEstado, 15) & ImpreFormat(R!cMonedaC, 8) & ImpreFormat(R!nMontoCol, 12, 2) & ImpreFormat(Format(R!dVigencia, "dd-mm-yyyy"), 12) & ImpreFormat(R!cUser, 8) & lnSaltoLinDoc
            sCadTemp = sCadTemp & Space(6) & ImpreFormat("COD.CLIENTE", 15) & ImpreFormat("COD.GARANT.", 12) & ImpreFormat("MONEDA", 8) & ImpreFormat("TIPO GARANT.", 12) & ImpreFormat("CLASE GARANT.", 14) & ImpreFormat("MONTO UTI. MN", 14) & ImpreFormat("MONTO UTI. ME", 13) & lnSaltoLinDoc
        Else
            sCadTemp = sCadTemp & Space(6) & ImpreFormat(R!cPersCodG, 15) & ImpreFormat(R!cNumGarant, 12) & ImpreFormat(R!cMonedaG, 8) & ImpreFormat(R!sSubTipo, 12) & ImpreFormat(R!cTipo, 13) & ImpreFormat(R!nMontoSoles, 13, 2) & ImpreFormat(R!nMontoDolares, 13, 2) & lnSaltoLinDoc
            nTotalSoles = nTotalSoles + R!nMontoSoles
            nTotalDolares = nTotalDolares + R!nMontoDolares

        End If
        If P Mod 500 = 0 Then
            sCadImp = sCadImp & sCadTemp
            sCadTemp = ""
         End If
        R.MoveNext
        
                
    Loop
    
    sCadImp = sCadImp & lnSaltoLinDoc & ImpreFormat("T O T A L  R E P O R T E:", 78) & ImpreFormat(nTotalSoles, 13, 2) & ImpreFormat(nTotalDolares, 11, 2) & lnSaltoLinDoc
    
    R.Close
    Set R = Nothing

    ImprimeGarantiasXClientes = sCadImp

End Function

Public Function ImprimeReporteGarantes(ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatAgencias As Variant, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaGarantiasXGarantes(pMatAgencias)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE DE GARANTES AL: " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(4) & ImpreFormat("CUENTA", 18) & ImpreFormat("F.DESEMB.", 10) & ImpreFormat("ANALIST.", 8) & ImpreFormat("MONTO DESEMB.", 13) & ImpreFormat("MONEDA", 8) & ImpreFormat("COD.GARANT.", 11) & ImpreFormat("TIPO GARANT.", 14) & ImpreFormat("MONTO GARANT.", 13) & ImpreFormat("MONTO GRAVADO", 13) & lnSaltoLinDoc
        sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc
        sCadCab = sCadImp
        sCadImp = sCadImp & lsTab & "MONEDA : " & Trim(R!sMoneda) & lnSaltoLinDoc
        sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "Garante: " & ImpreFormat(R!cPersCod, 15) & " " & ImpreFormat(R!sGarante, 30) & lnSaltoLinDoc
        'R.MoveNext
    End If
    
    Do While Not R.EOF
        If Trim(R!sGarantia) = "ZTOTAL" Then

            R.MoveNext
            If R.EOF Then Exit Do
                If Trim(R!sMoneda) <> "ZTOTAL" Then
                    If Trim(R!sGarante) <> "ZTOTAL" Then
                        sCadImp = sCadImp & lsTab & "Garante: " & ImpreFormat(R!cPersCod, 15) & " " & ImpreFormat(R!sGarante, 30) & lnSaltoLinDoc
                    End If
                Else
                    R.MoveNext
                    If R.EOF Then Exit Do
                End If

                'TOTAL FINAL
                If Trim(R!sGarante) = "ZTOTAL" Then
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "TOTAL MONEDA: " & ImpreFormat("", 85) & ImpreFormat(R!nMontoGarantia, 12) & ImpreFormat(R!nMontoGarantizado, 12) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then Exit Do
                    If Trim(R!sMoneda) <> "ZTOTAL" Then
                        sCadImp = sCadImp & lsTab & "MONEDA : " & Trim(R!sMoneda) & lnSaltoLinDoc
                        sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "Garante: " & ImpreFormat(R!cPersCod, 15) & " " & ImpreFormat(R!sGarante, 30) & lnSaltoLinDoc
                    Else
                        Exit Do
                    End If
                End If

        End If
    
        sCadImp = sCadImp & Space(4) & ImpreFormat(R!cCtaCod, 18) & ImpreFormat(Format(R!dDesembolso, "dd-mm-yyyy"), 10) & ImpreFormat(R!sAnalista, 8) & ImpreFormat(R!nMontoCol, 13) & ImpreFormat(R!cMoneda, 8) & ImpreFormat(R!sGarantia, 11) & ImpreFormat(R!cTipo, 12) & ImpreFormat(R!nMontoGarantia, 13, 2) & ImpreFormat(R!nMontoGarantizado, 13, 2) & lnSaltoLinDoc
                                    
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeReporteGarantes = sCadImp

End Function

Public Function ImprimeGarantiasXTipoDetallado(ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatAgencias As Variant, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaGarantiasXTipoDetallado(pMatAgencias)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE DE GARANTIAS POR TIPO - DETALLADO AL: " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(2) & ImpreFormat("CUENTA", 18) & ImpreFormat("COD.CLIENTE", 15) & ImpreFormat("NOMBRE CLIENTE", 30) & ImpreFormat("FECHA DESEMB.", 13) & ImpreFormat("MONEDA", 8) & ImpreFormat("MONTO DESEMB.", 13) & lnSaltoLinDoc
        sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc
        sCadCab = sCadImp
        sCadImp = sCadImp & lsTab & "CLASE DE GARANTIA : " & Trim(R!sTipo) & lnSaltoLinDoc
        sCadImp = sCadImp & String(140, "=") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "MONEDA DE GARANTIA : " & Trim(R!sMoneda) & lnSaltoLinDoc
        sCadImp = sCadImp & String(140, "*") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "TIPO DE GARANTIA : " & Trim(R!sSubTipo) & lnSaltoLinDoc
        'sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc
        'sCadImp = sCadImp & Space(6) & ImpreFormat("COD.GARANTIA:", 20) & ImpreFormat(R!sGarantia, 8) & ImpreFormat("FECHA REGISTRO:", 20) & ImpreFormat(Format(R!DGarantia, "dd-mm-yyyy"), 10) & ImpreFormat("MONTO GARANTIA:", 20) & ImpreFormat(R!nMontoGarantia, 10, 2) & lnSaltoLinDoc
        'R.MoveNext
    End If
    
    Do While Not R.EOF
    'FINAL DE GARANTIA
        If Trim(R!sGarantia) = "ZTOTAL" Then
            sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("TOTAL TIPO DE GARANTIA", 120) & ImpreFormat(R!nMontoGarantia, 11, 2) & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE TIPO
            If R!sSubTipo = "ZTOTAL" Then
                sCadImp = sCadImp & lsTab & String(140, "*") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("TOTAL MONEDA", 120) & ImpreFormat(R!nMontoGarantia, 11, 2) & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL MONEDA
                If R!sMoneda = "ZTOTAL" Then
                    sCadImp = sCadImp & lsTab & String(140, "=") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("TOTAL CLASE GARANTIA", 120) & ImpreFormat(R!nMontoGarantia, 11, 2) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then Exit Do
                    
                    'FINAL TIPO
                    If R!sTipo = "ZTOTAL" Then
                        sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
                        'sCadImp = sCadImp & lsTab & ImpreFormat("TOTAL CLASE GARANTIA", 15) & ImpreFormat(R!nMontoGarantia, 11, 2) & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        R.MoveNext
                        If R.EOF Then Exit Do
                    Else
                        nPuntPag = 0
                        sCadImp = sCadImp & Chr$(12)
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & sCadCab
                        sCadImp = sCadImp & lsTab & "CLASE DE GARANTIA : " & Trim(R!sTipo) & lnSaltoLinDoc
                        sCadImp = sCadImp & String(140, "=") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "MONEDA DE GARANTIA : " & Trim(R!sMoneda) & lnSaltoLinDoc
                        sCadImp = sCadImp & String(140, "*") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "TIPO DE GARANTIA : " & Trim(R!sSubTipo) & lnSaltoLinDoc
                        sCadImp = sCadImp & String(140, ">") & lnSaltoLinDoc
                    End If
                Else
                
                    'sCadImp = sCadImp & lsTab & "CLASE DE GARANTIA : " & Trim(R!sTipo) & lnSaltoLinDoc
                    'sCadImp = sCadImp & String(140, "=") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "MONEDA DE GARANTIA : " & Trim(R!sMoneda) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(140, "*") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "TIPO DE GARANTIA : " & Trim(R!sSubTipo) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(140, ">") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "TIPO DE GARANTIA : " & Trim(R!sSubTipo) & lnSaltoLinDoc
                sCadImp = sCadImp & String(140, ">") & lnSaltoLinDoc
            End If
        End If
        
        If R!cCtaCod = "ZTOTAL" Then
        '    sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc
        '    sCadImp = sCadImp & Space(6) & ImpreFormat("COD.GARANTIA:", 20) & ImpreFormat(R!sGarantia, 8) & ImpreFormat("FECHA REGISTRO:", 20) & ImpreFormat(Format(R!DGarantia, "dd-mm-yyyy"), 10) & ImpreFormat("MONTO GARANTIA:", 20) & ImpreFormat(R!nMontoGarantia, 10, 2) & lnSaltoLinDoc
        Else
            sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & Space(6) & ImpreFormat("COD.GARANTIA:", 20) & ImpreFormat(R!sGarantia, 8) & Space(15) & ImpreFormat("FECHA REGISTRO:", 15) & ImpreFormat(Format(R!DGarantia, "dd-mm-yyyy"), 10) & Space(20) & ImpreFormat("MONTO GARANTIA:", 18) & ImpreFormat(R!nMontoGarantia, 10, 2) & lnSaltoLinDoc

            sCadImp = sCadImp & Space(2) & ImpreFormat(R!cCtaCod, 18) & ImpreFormat(R!cPersCod, 15) & ImpreFormat(R!cpersnombre, 30) & ImpreFormat(Format(R!dVigencia, "dd-mm-yyyy"), 13) & ImpreFormat(R!cMoneda, 8) & ImpreFormat(R!nMontoCol, 11, 2) & lnSaltoLinDoc
        End If
        R.MoveNext
    Loop
    R.Close
    
    Set R = Nothing
    ImprimeGarantiasXTipoDetallado = sCadImp

End Function

Public Function ImprimeGarantiasXTipoResumido(ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, ByVal pMatAgencias As Variant, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaGarantiasXTipoResumido(pMatAgencias)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE DE GARANTIAS POR TIPO - RESUMIDO AL: " & Format(pdFecSis, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(140, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(45) & ImpreFormat("TIPO GARANTIA", 45) & ImpreFormat("MONTO GARANTIA", 15) & lnSaltoLinDoc
        sCadImp = sCadImp & String(140, "-") & lnSaltoLinDoc
        sCadCab = sCadImp
        sCadImp = sCadImp & lsTab & "CLASE DE GARANTIA : " & Trim(R!sTipo) & lnSaltoLinDoc
        sCadImp = sCadImp & String(140, "=") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE GARANTIA
        If Trim(R!sSubTipo) = "ZTOTAL" Then
            sCadImp = sCadImp & Space(45) & ImpreFormat("TOTAL " & R!sMoneda, 25) & ImpreFormat(R!nMontoGarantia, 15, 2) & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE TIPO
            If R!sMoneda = "ZTOTAL" And R!sSubTipo = "ZTOTAL" Then
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                If R!sTipo = "ZTOTAL" Then
                    Exit Do
                Else
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "CLASE DE GARANTIA : " & Trim(R!sTipo) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(140, "=") & lnSaltoLinDoc
                End If
            End If
        End If
        
        sCadImp = sCadImp & Space(38) & ImpreFormat("TOTAL " & R!sSubTipo, 45) & ImpreFormat(R!nMontoGarantia, 15, 2) & lnSaltoLinDoc
        
        R.MoveNext
    Loop
    R.Close
    
    Set R = Nothing
    ImprimeGarantiasXTipoResumido = sCadImp

End Function

Public Function ImprimeListaReprogramados(ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal psNomAge As String, ByVal pMatAgencias As Variant, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String
Dim nMontoDesemb As Double
Dim nCapitalPag As Double

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaDatosListaReprogramados(pMatAgencias, pdFecIni, pdFecFin)
    Set oDCred = Nothing
    
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "REPORTE DE REFINANCIADOS DEL : " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(112, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(2) & ImpreFormat("CUENTA", 18) & ImpreFormat("CLIENTE", 30) & ImpreFormat("F.VIGENCIA", 11) & ImpreFormat("MONTO DESEMB.", 14) & ImpreFormat("CAPITAL PAG.", 15) & ImpreFormat("SALDO CAP.", 13) & lnSaltoLinDoc
        sCadImp = sCadImp & String(112, "-") & lnSaltoLinDoc
        sCadCab = sCadImp
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
        sCadImp = sCadImp & String(112, "=") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "MONEDA : " & Trim(R!sMoneda) & lnSaltoLinDoc
        sCadImp = sCadImp & String(112, "*") & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!sAnalista) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(112, "-") & lnSaltoLinDoc
    End If
    
    Do While Not R.EOF
    'FINAL DE ANALISTA
        If Trim(R!cCtaCod) = "TOTAL" Then
            sCadImp = sCadImp & lsTab & String(112, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & ImpreFormat("TOTAL " & R!sAnalista, 64) & ImpreFormat(R!nMontoDesemb, 12, 2) & ImpreFormat(R!nCapPagado, 12, 2) & ImpreFormat(R!nSaldoCap, 12, 2) & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            R.MoveNext
            If R.EOF Then
                Exit Do
            End If
            
            'FINAL DE TIPO
            If R!sAnalista = "TOTAL" Then
                sCadImp = sCadImp & lsTab & String(112, "*") & lnSaltoLinDoc
                sCadImp = sCadImp & lsTab & ImpreFormat("TOTAL " & R!sMoneda, 64) & ImpreFormat(R!nMontoDesemb, 12, 2) & ImpreFormat(R!nCapPagado, 12, 2) & ImpreFormat(R!nSaldoCap, 12, 2) & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL MONEDA
                If R!sMoneda = "TOTAL" Then
                    sCadImp = sCadImp & lsTab & String(112, "=") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & ImpreFormat("TOTAL " & R!sAgencia, 64) & ImpreFormat(R!nMontoDesemb, 12, 2) & ImpreFormat(R!nCapPagado, 12, 2) & ImpreFormat(R!nSaldoCap, 12, 2) & lnSaltoLinDoc
                    sCadImp = sCadImp & lnSaltoLinDoc
                    R.MoveNext
                    If R.EOF Then Exit Do
                    
                    If R!sAgencia = "TOTAL" Then
                        Exit Do
                    Else
                        nPuntPag = 0
                        sCadImp = sCadImp & Chr$(12)
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & lnSaltoLinDoc
                        sCadImp = sCadImp & sCadCab
                        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!sAgencia) & lnSaltoLinDoc
                        sCadImp = sCadImp & String(112, "=") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "MONEDA : " & Trim(R!sMoneda) & lnSaltoLinDoc
                        sCadImp = sCadImp & String(112, "*") & lnSaltoLinDoc
                        sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!sAnalista) & lnSaltoLinDoc
                    End If
                Else
                
                    sCadImp = sCadImp & lsTab & "MONEDA : " & Trim(R!sMoneda) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(112, "*") & lnSaltoLinDoc
                    sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!sAnalista) & lnSaltoLinDoc
                    sCadImp = sCadImp & String(112, "-") & lnSaltoLinDoc
                End If
            Else
                sCadImp = sCadImp & lsTab & "ANALISTA : " & Trim(R!sAnalista) & lnSaltoLinDoc
                sCadImp = sCadImp & String(112, "-") & lnSaltoLinDoc
            
            End If
        
        End If
        'sCadImp = sCadImp & Space(2) & ImpreFormat("CUENTA", 18) & ImpreFormat("CLIENTE", 30) & ImpreFormat("F.VIGENCIA", 11) & ImpreFormat("MONTO DESEMB.", 14) & ImpreFormat("CAPITAL PAG.", 13) & ImpreFormat("SALDO CAP.", 13) & lnSaltoLinDoc
        sCadImp = sCadImp & Space(2) & ImpreFormat(R!cCtaCod, 18) & ImpreFormat(R!sCliente, 30) & ImpreFormat(Format(R!dVigencia, "dd-mm-yyyy"), 11) & ImpreFormat(R!nMontoDesemb, 12, 2) & ImpreFormat(R!nCapPagado, 12, 2) & ImpreFormat(R!nSaldoCap, 12, 2) & lnSaltoLinDoc
        R.MoveNext
    Loop
    R.Close
    
    Set R = Nothing
    ImprimeListaReprogramados = sCadImp

End Function

'Creado por DAOR
Public Function nRepo108331_CreditosDesembolsadosBancoNacion(pdFecIni As Date, pdFecFin As Date, ByVal pMatAgencias As Variant, _
    ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "") As String
        
    Dim oDCred As DCredDoc
    Dim rs As ADODB.Recordset
    Dim sCadImp As String, sCadCabT As String
    Dim lsCadBuffer As String
    Dim lnIndice As Long
    Dim lnLineas As Integer, lnPage As Integer, I As Integer
    Dim p13 As Double, p46 As Double, p79 As Double, p1012 As Double, TriAdi As Double
    Dim nTEA As Double
    Dim nComision As Double
    Dim nTotalDes As Double, nTotalCom As Double
    
    'MAVM 20092808 *******  Doc Desarrollo Item 23
    Dim oITF As COMDConstSistema.FCOMITF
    Dim lnValor As Double
    Set oITF = New COMDConstSistema.FCOMITF
    oITF.fgITFParametros
    'Fin MAVM 20092808 *******
    
'On Error GoTo dError

    Set oDCred = New DCredDoc
    Set rs = oDCred.RecuperaCreditosDesembolsadosEnBancoNacion(pdFecIni, pdFecFin, pMatAgencias)
    Set oDCred = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        lnLineas = 0
        lnPage = 1
        
        'MAVM 29052009 Se modifico los valores segun Memorandum 20-2009 CR-CG/CMACM
        p13 = 0.017 '0.02
        p46 = 0.022 '0.025
        p79 = 0.027 '0.03
        p1012 = 0.032 '0.035
        TriAdi = 0.005
        nTotalDes = 0
        nTotalCom = 0
        'End MAVM

        'CABECERA
        sCadCab = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS DESEMBOLSADOS EN AGENCIAS DEL BANCO DE LA NACIÓN", psNomCmac, , , "108331", "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        sCadCabT = sCadCabT & String(238, "-") & lnSaltoLinDoc
        sCadCabT = sCadCabT & Space(1) & ImpreFormat("Nº", 4)
        sCadCabT = sCadCabT & ImpreFormat("Ofi. Bco. Nac.", 15)
        sCadCabT = sCadCabT & ImpreFormat("Agencia", 15)
        sCadCabT = sCadCabT & ImpreFormat("Nº Credito", 18)
        sCadCabT = sCadCabT & ImpreFormat("Cod. Cliente", 13)
        sCadCabT = sCadCabT & ImpreFormat("Cliente", 25)
        sCadCabT = sCadCabT & ImpreFormat("Linea", 15)
        sCadCabT = sCadCabT & ImpreFormat("Fec. Aprob.", 10)
        sCadCabT = sCadCabT & ImpreFormat("Fec. Desem.", 10)
        sCadCabT = sCadCabT & ImpreFormat("Desembolso", 10)
        sCadCabT = sCadCabT & ImpreFormat("Monto Pagado", 10)
        sCadCabT = sCadCabT & ImpreFormat("Mone.", 5)
        sCadCabT = sCadCabT & ImpreFormat("Cuotas", 6)
        sCadCabT = sCadCabT & ImpreFormat("Comisión", 10)
        sCadCabT = sCadCabT & ImpreFormat("T.E.M", 6)
        sCadCabT = sCadCabT & ImpreFormat("T.E.A", 5)
        sCadCabT = sCadCabT & ImpreFormat("Gracia", 6)
        sCadCabT = sCadCabT & ImpreFormat("Fec. Cuota", 10)
        sCadCabT = sCadCabT & ImpreFormat("Gar.Real", 8) & lnSaltoLinDoc
        sCadCabT = sCadCabT & String(238, "-") & lnSaltoLinDoc
        
        sCadCab = sCadImp + sCadCabT
        sCadImp = sCadImp + sCadCabT
        lnIndice = 0
        lnLineas = 7
        I = 0
        Do While Not rs.EOF
            I = I + 1
            lnIndice = lnIndice + 1
            lnLineas = lnLineas + 1
            
            nComision = 0
            'Cadena de Impresion
            sCadImp = sCadImp & ImpreFormat(I, 4)
            sCadImp = sCadImp & ImpreFormat(rs!vAgencia, 15)
            sCadImp = sCadImp & ImpreFormat(rs!cAgeDescripcion, 15)
            sCadImp = sCadImp & ImpreFormat(rs!cCtaCod, 18)
            sCadImp = sCadImp & ImpreFormat(rs!cPersCod, 13)
            sCadImp = sCadImp & ImpreFormat(rs!cpersnombre, 30)
            sCadImp = sCadImp & ImpreFormat(rs!LineaCredito, 20)
            sCadImp = sCadImp & ImpreFormat(Format(rs!FechaAprob, "dd/mm/yyyy"), 10)
            sCadImp = sCadImp & ImpreFormat(Format(rs!FechaDesem, "dd/mm/yyyy"), 10)
            
            'MAVM 20092808 *******
            lnValor = 0
            lnValor = rs!nMonto * oITF.gnITFPorcent
            lnValor = oITF.CortaDosITF(lnValor)
            sCadImp = sCadImp & ImpreFormat(rs!nMonto, 9, 2, True)
            sCadImp = sCadImp & ImpreFormat(rs!nMonto - (lnValor), 9, 2, True)
            'MAVM 20092808 *******
            
            sCadImp = sCadImp & ImpreFormat(IIf(rs!Moneda = 1, "MN", "ME"), 4)
            sCadImp = sCadImp & ImpreFormat(rs!nCuotas, 6)
'            nComision = IIf(rs!nCuotas <= 3, rs!nMonto * p13, _
'                            IIf(rs!nCuotas <= 6, rs!nMonto * p46, _
'                                IIf(rs!nCuotas <= 9, rs!nMonto * p79, _
'                                    IIf(rs!nCuotas <= 12, rs!nMonto * p1012, _
'                                        IIf((rs!nCuotas Mod 3) = 0, _
'                                            rs!nMonto * (((rs!nCuotas - 12) / 3 * TriAdi) + p1012), _
'                                            rs!nMonto * ((((rs!nCuotas - (rs!nCuotas Mod 3) - 12) / 3) * TriAdi) + p1012))))))

'            nComision = IIf(rs!nCuotas <= 3, rs!nMonto * p13, _
'                            IIf(rs!nCuotas <= 6, rs!nMonto * p46, _
'                                IIf(rs!nCuotas <= 9, rs!nMonto * p79, _
'                                    IIf(rs!nCuotas <= 12, rs!nMonto * p1012, rs!nMonto * ((((rs!nCuotas - 12) \ 3) * TriAdi) + p1012) _
'                                        ))))
            'MAVM 20091006
            nComision = IIf(rs!nCuotas <= 3, (rs!nMonto - (lnValor)) * p13, _
                            IIf(rs!nCuotas <= 6, (rs!nMonto - (lnValor)) * p46, _
                                IIf(rs!nCuotas <= 9, (rs!nMonto - (lnValor)) * p79, _
                                    IIf(rs!nCuotas <= 12, (rs!nMonto - (lnValor)) * p1012, _
                                        IIf((rs!nCuotas Mod 3) = 0, _
                                            (rs!nMonto - (lnValor)) * (((rs!nCuotas - 12) / 3 * TriAdi) + p1012), _
                                            (rs!nMonto - (lnValor)) * (((((rs!nCuotas - (rs!nCuotas Mod 3) - 12) / 3) + 1) * TriAdi) + p1012))))))
                                        
            sCadImp = sCadImp & ImpreFormat(Round(nComision, 2), 9, 2, True)
            sCadImp = sCadImp & ImpreFormat(rs!nTasaInteres, 4, 2) & "%"
            nTEA = ((1 + rs!nTasaInteres / 100) ^ 12 - 1) * 100
            sCadImp = sCadImp & ImpreFormat(nTEA, 4, 2) & "%"
            sCadImp = sCadImp & ImpreFormat(rs!nPeriodoGracia, 4)
            sCadImp = sCadImp & ImpreFormat(Format(rs!FechaPago, "dd/mm/yyyy"), 10)
            sCadImp = sCadImp & ImpreFormat(rs!GarantiaReal, 8)
            sCadImp = sCadImp & lnSaltoLinDoc
            nTotalDes = nTotalDes + rs!nMonto
            nTotalCom = nTotalCom + nComision
            
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & sCadImp
                sCadImp = ""
            End If
            
'            If lnLineas >= 55 Then
'                lnPage = lnPage + 1
'                lnLineas = 7
'                sCadImp = sCadImp & Chr$(12)
'                sCadImp = sCadImp & nImprimeCabeceraReportes(psNomCmac, psNomAge, psCodUser, pdFecSis, "CREDITOS DESEMBOLSADOS EN AGENCIAS DEL BANCO DE LA NACIÓN", _
'                "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"), lnPage, 150, "", "108331")
'                sCadImp = sCadImp & String(238, "-") & lnSaltoLinDoc
'                sCadImp = sCadImp & Space(1) & ImpreFormat("Nº", 4)
'                sCadImp = sCadImp & ImpreFormat("Ofi. Bco. Nac.", 15)
'                sCadImp = sCadImp & ImpreFormat("Agencia", 15)
'                sCadImp = sCadImp & ImpreFormat("Nº Credito", 18)
'                sCadImp = sCadImp & ImpreFormat("Cod. Cliente", 13)
'                sCadImp = sCadImp & ImpreFormat("Cliente", 30)
'                sCadImp = sCadImp & ImpreFormat("Linea", 20)
'                sCadImp = sCadImp & ImpreFormat("Fec. Aprob.", 10)
'                sCadImp = sCadImp & ImpreFormat("Fec. Desem.", 10)
'                sCadImp = sCadImp & ImpreFormat("Monto", 10)
'                sCadImp = sCadImp & ImpreFormat("Mone.", 5)
'                sCadImp = sCadImp & ImpreFormat("Cuotas", 6)
'                sCadImp = sCadImp & ImpreFormat("Comisión", 10)
'                sCadImp = sCadImp & ImpreFormat("T.E.M", 6)
'                sCadImp = sCadImp & ImpreFormat("T.E.A", 5)
'                sCadImp = sCadImp & ImpreFormat("Gracia", 6)
'                sCadImp = sCadImp & ImpreFormat("Fec. Cuota", 10)
'                sCadImp = sCadImp & ImpreFormat("Gar.Real", 8) & lnSaltoLinDoc
'                sCadImp = sCadImp & String(238, "-") & lnSaltoLinDoc
'            End If
            rs.MoveNext
        Loop
        sCadImp = sCadImp & String(238, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(140) & ImpreFormat("Total : ", 12) & ImpreFormat(nTotalDes, 9, 2, True) & Space(15) & ImpreFormat(nTotalCom, 9, 2, True) & lnSaltoLinDoc
       
        lsCadBuffer = lsCadBuffer & sCadImp & lnSaltoLinDoc
        nRepo108331_CreditosDesembolsadosBancoNacion = lsCadBuffer
    Else
        nRepo108331_CreditosDesembolsadosBancoNacion = ""
    End If
End Function

'Creado por DAOR
Public Function nRepo108331_CreditosDesembolsadosBancoNacionExcel(pdFecIni As Date, pdFecFin As Date, ByVal pMatAgencias As Variant, _
    ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "") As String

    Dim oDCred As DCredDoc
    Dim rs As ADODB.Recordset
    Dim lbExisteHoja  As Boolean
    Dim lsNomHoja As String
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim glsArchivo As String
    Dim liLineas As Integer, I As Integer
    Dim nNum As Double
    Dim fs As Scripting.FileSystemObject
    Dim p13 As Double, p46 As Double, p79 As Double, p1012 As Double, TriAdi As Double
    Dim nTEA As Double
    Dim nComision As Double
    Dim nTotalDes As Double, nTotalCom As Double
    
    'MAVM 20092808 *******  Doc Desarrollo Item 23
    Dim oITF As COMDConstSistema.FCOMITF
    Dim lnValor As Double
    Set oITF = New COMDConstSistema.FCOMITF
    oITF.fgITFParametros
    'Fin MAVM 20092808 *******
    
    Set oDCred = New DCredDoc
    Set rs = oDCred.RecuperaCreditosDesembolsadosEnBancoNacion(pdFecIni, pdFecFin, pMatAgencias)
    Set oDCred = Nothing

    If Not (rs.EOF And rs.BOF) Then
        glsArchivo = "Creditos_Desemb_Bco_Nac" & Format(pdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    
        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
            Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
        Else
            Set xlLibro = xlAplicacion.Workbooks.Add
        End If
        
        Set xlHoja1 = xlLibro.Worksheets.Add
        lbExisteHoja = False
        lsNomHoja = "Creditos_Desemb_Bco_Nac"
        For Each xlHoja1 In xlLibro.Worksheets
            If xlHoja1.Name = lsNomHoja Then
                xlHoja1.Activate
                lbExisteHoja = True
                Exit For
            End If
        Next
        If lbExisteHoja = False Then
            Set xlHoja1 = xlLibro.Worksheets.Add
            xlHoja1.Name = lsNomHoja
        End If
    
        xlAplicacion.Range("A1:A1").ColumnWidth = 3
        xlAplicacion.Range("B1:B1").ColumnWidth = 10
        xlAplicacion.Range("C1:C1").ColumnWidth = 10
        xlAplicacion.Range("D1:D1").ColumnWidth = 12
        xlAplicacion.Range("E1:E1").ColumnWidth = 12
        xlAplicacion.Range("F1:F1").ColumnWidth = 20
        xlAplicacion.Range("G1:G1").ColumnWidth = 15
        xlAplicacion.Range("H1:H1").ColumnWidth = 10
        xlAplicacion.Range("I1:I1").ColumnWidth = 10
        xlAplicacion.Range("J1:J1").ColumnWidth = 10
        'MAVM 20090901
        xlAplicacion.Range("K1:K1").ColumnWidth = 10
        'End MAVM 20090901
        xlAplicacion.Range("L1:L1").ColumnWidth = 5
        xlAplicacion.Range("M1:M1").ColumnWidth = 5
        xlAplicacion.Range("N1:N1").ColumnWidth = 10
        xlAplicacion.Range("O1:O1").ColumnWidth = 5
        xlAplicacion.Range("P1:P1").ColumnWidth = 7
        xlAplicacion.Range("Q1:Q1").ColumnWidth = 5
        xlAplicacion.Range("R1:R1").ColumnWidth = 10
        xlAplicacion.Range("S1:S1").ColumnWidth = 5
        xlAplicacion.Range("A1:Z2000").Font.Size = 8
        
        xlHoja1.Cells(1, 1) = psNomCmac
        xlHoja1.Cells(2, 1) = psNomAge
        xlHoja1.Cells(1, 17) = Trim(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"))
        xlHoja1.Cells(2, 17) = psCodUser
                
        xlHoja1.Cells(4, 1) = "CREDITOS DESEMBOLSADOS EN AGENCIAS DEL BANCO DE LA NACIÓN"
        xlHoja1.Cells(5, 1) = "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy")
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, 19)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 19)).Merge True
        xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 19)).Merge True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, 19)).HorizontalAlignment = xlCenter
               
        liLineas = 6
        xlHoja1.Cells(liLineas, 1) = "Nº"
        xlHoja1.Cells(liLineas, 2) = "Ofi. Bco. Nac"
        xlHoja1.Cells(liLineas, 3) = "Agencia"
        xlHoja1.Cells(liLineas, 4) = "Nº Credito"
        xlHoja1.Cells(liLineas, 5) = "Cod. Cliente"
        xlHoja1.Cells(liLineas, 6) = "Cliente"
        xlHoja1.Cells(liLineas, 7) = "Linea"
        xlHoja1.Cells(liLineas, 8) = "Fec. Aproba."
        xlHoja1.Cells(liLineas, 9) = "Fec. Desemb."
        xlHoja1.Cells(liLineas, 10) = "Desembolso"
        'MAVM 20090901
        xlHoja1.Cells(liLineas, 11) = "Monto Pagado"
        'MAVM 20090901
        xlHoja1.Cells(liLineas, 12) = "Moneda"
        xlHoja1.Cells(liLineas, 13) = "Cuotas"
        xlHoja1.Cells(liLineas, 14) = "Comisión"
        xlHoja1.Cells(liLineas, 15) = "T.E.M"
        xlHoja1.Cells(liLineas, 16) = "T.E.A"
        xlHoja1.Cells(liLineas, 17) = "Gracia"
        xlHoja1.Cells(liLineas, 18) = "Fec. 1ra. Cuota"
        xlHoja1.Cells(liLineas, 19) = "Garantías Reales"
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 19)).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 19)).HorizontalAlignment = xlCenter
        
        
        'MAVM 29052009 Se modifico los valores segun Memorandum 20-2009 CR-CG/CMACM
        p13 = 0.017 '0.02
        p46 = 0.022 '0.025
        p79 = 0.027 '0.03
        p1012 = 0.032 '0.035
        TriAdi = 0.005
        nTotalDes = 0
        nTotalCom = 0
        'End MAVM
        
        liLineas = liLineas + 1
        I = 0
        rs.MoveFirst
        Do Until rs.EOF
            I = I + 1
            nComision = 0
            xlHoja1.Cells(liLineas, 1) = I
            xlHoja1.Cells(liLineas, 2) = rs!vAgencia
            xlHoja1.Cells(liLineas, 3) = rs!cAgeDescripcion
            xlHoja1.Cells(liLineas, 4) = rs!cCtaCod
            xlHoja1.Cells(liLineas, 5) = rs!cPersCod
            xlHoja1.Cells(liLineas, 6) = rs!cpersnombre
            xlHoja1.Cells(liLineas, 7) = rs!LineaCredito
            xlHoja1.Cells(liLineas, 8) = Format(rs!FechaAprob, "dd/mm/yyyy")
            xlHoja1.Cells(liLineas, 9) = Format(rs!FechaDesem, "dd/mm/yyyy")
            xlHoja1.Cells(liLineas, 10) = ImpreFormat(rs!nMonto, 9, 2, True)
            
            'MAVM 20092808 *******
            lnValor = 0
            lnValor = rs!nMonto * oITF.gnITFPorcent
            xlHoja1.Cells(liLineas, 11) = ImpreFormat((rs!nMonto - (lnValor)), 9, 2, True)
            'End MAVM 20092808 *******
                    
            xlHoja1.Cells(liLineas, 12) = IIf(rs!Moneda = 1, "MN", "ME")
            xlHoja1.Cells(liLineas, 13) = rs!nCuotas
'            nComision = IIf(rs!nCuotas <= 3, rs!nMonto * p13, _
'                            IIf(rs!nCuotas <= 6, rs!nMonto * p46, _
'                                IIf(rs!nCuotas <= 9, rs!nMonto * p79, _
'                                    IIf(rs!nCuotas <= 12, rs!nMonto * p1012, _
'                                        IIf((rs!nCuotas Mod 3) = 0, _
'                                            rs!nMonto * (((rs!nCuotas - 12) / 3 * TriAdi) + p1012), _
'                                            rs!nMonto * ((((rs!nCuotas - (rs!nCuotas Mod 3) - 12) / 3) * TriAdi) + p1012))))))
            
'            nComision = IIf(rs!nCuotas <= 3, rs!nMonto * p13, _
'                            IIf(rs!nCuotas <= 6, rs!nMonto * p46, _
'                                IIf(rs!nCuotas <= 9, rs!nMonto * p79, _
'                                    IIf(rs!nCuotas <= 12, rs!nMonto * p1012, rs!nMonto * ((((rs!nCuotas - 12) \ 3) * TriAdi) + p1012) _
'                                        ))))
            'MAVM 20091006
            nComision = IIf(rs!nCuotas <= 3, (rs!nMonto - (lnValor)) * p13, _
                            IIf(rs!nCuotas <= 6, (rs!nMonto - (lnValor)) * p46, _
                                IIf(rs!nCuotas <= 9, (rs!nMonto - (lnValor)) * p79, _
                                    IIf(rs!nCuotas <= 12, (rs!nMonto - (lnValor)) * p1012, _
                                        IIf((rs!nCuotas Mod 3) = 0, _
                                            (rs!nMonto - (lnValor)) * (((rs!nCuotas - 12) / 3 * TriAdi) + p1012), _
                                            (rs!nMonto - (lnValor)) * (((((rs!nCuotas - (rs!nCuotas Mod 3) - 12) / 3) + 1) * TriAdi) + p1012))))))
            
            xlHoja1.Cells(liLineas, 14) = ImpreFormat(Round(nComision, 2), 9, 2, True)
            xlHoja1.Cells(liLineas, 15) = ImpreFormat(rs!nTasaInteres, 9, 2) & "%"
            nTEA = ((1 + rs!nTasaInteres / 100) ^ 12 - 1) * 100
            xlHoja1.Cells(liLineas, 16) = ImpreFormat(nTEA, 9, 2) & "%"
            xlHoja1.Cells(liLineas, 17) = rs!nPeriodoGracia
            xlHoja1.Cells(liLineas, 18) = Format(rs!FechaPago, "dd/mm/yyyy")
            xlHoja1.Cells(liLineas, 19) = rs!GarantiaReal
            xlHoja1.Range(xlHoja1.Cells(liLineas - 1, 1), xlHoja1.Cells(liLineas, 19)).Borders.LineStyle = 1
            nTotalDes = nTotalDes + rs!nMonto
            nTotalCom = nTotalCom + nComision
            rs.MoveNext
            liLineas = liLineas + 1
        Loop
        xlHoja1.Cells(liLineas, 9) = "TOTAL"
        xlHoja1.Cells(liLineas, 10) = ImpreFormat(nTotalDes, 9, 2, True)
        xlHoja1.Cells(liLineas, 13) = ImpreFormat(nTotalCom, 9, 2, True)
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        xlAplicacion.Visible = True
        xlAplicacion.Windows(1).Visible = True
        'xlLibro.Close 'Cierra el libro de trabajo
        'xlAplicacion.Quit 'Cierra Microsoft Excel con el método Quit.
        'Libera los objetos.
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        nRepo108331_CreditosDesembolsadosBancoNacionExcel = ""
     Else
        MsgBox "No existen datos para generar el reporte"
        nRepo108331_CreditosDesembolsadosBancoNacionExcel = ""
     End If
End Function

Public Function ImprimeListadoCreditosReprogramados(ByVal psCodUser As String, ByVal pdFecSis As Date, _
                    ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
                    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "") As String

Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String
Dim I As Integer
Dim sAgencia As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaDatosCreditosReprogramados(pMatAgencias, pdFecIni, pdFecFin)
    Set oDCred = Nothing
    
    sCadImp = ""
    
    If R.RecordCount > 0 Then
        R.MoveFirst
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS REPROGRAMADOS DEL: " & Format(pdFecIni, "dd/mm/yyyy") & " AL:" & Format(pdFecFin, "dd/mm/yyyy"), psNomCmac, , True)
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadCab = sCadImp
        sAgencia = Trim(R!cAgeCod)
        sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!cAgeCod) & ": " & Trim(R!cAgeDescripcion) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(148, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("CREDITO", 20) & ImpreFormat("FECHA VIG.", 10) & ImpreFormat("CLIENTE", 34) & ImpreFormat("S.CAP.", 11) & ImpreFormat("MON.DESEM", 11) & ImpreFormat("NºCUOT.", 7) & ImpreFormat("CUO.REP.", 8) & ImpreFormat("ANALISTA", 12) & ImpreFormat("MONEDA", 7) & ImpreFormat("FECHAREPRO", 10) & lnSaltoLinDoc
        sCadImp = sCadImp & lsTab & String(148, "-") & lnSaltoLinDoc
        'R.MoveNext
    End If

    Do While Not R.EOF

        If sAgencia <> Trim(R!cAgeCod) Then
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & "AGENCIA : " & Trim(R!cAgeCod) & ": " & Trim(R!cAgeDescripcion) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(148, "-") & lnSaltoLinDoc
            sCadImp = sCadImp & ImpreFormat("CREDITO", 20) & ImpreFormat("FECHA VIG.", 10) & ImpreFormat("CLIENTE", 30) & ImpreFormat("S.CAP.", 38) & ImpreFormat("MON.DESEM", 11) & ImpreFormat("NºCUOT", 7) & ImpreFormat("NºCUO.REPRO.", 12) & ImpreFormat("ANALISTA", 8) & ImpreFormat("MONEDA", 7) & ImpreFormat("FECHAREPRO", 10) & lnSaltoLinDoc
            sCadImp = sCadImp & lsTab & String(148, "-") & lnSaltoLinDoc
        End If
        '***ALPA**20060605*************************************************************************************************************************************************
        sCadImp = sCadImp & ImpreFormat(R!cCtaCod, 20) & ImpreFormat(Format(R!dVigencia, "dd/mm/yyyy"), 10) & ImpreFormat(R!cpersnombre, 30) & ImpreFormat(R!nSaldoCap, 10) & ImpreFormat(R!nMonCol, 10) & ImpreFormat(R!nCuotas, 8, 0) & ImpreFormat(R!nCuota, 8, 0) & Space(8) & R!cUser & Space(10) & ImpreFormat(R!cMoneda, 8) & Mid(R!dFechaRepro, 1, 8) & lnSaltoLinDoc
        '******************************************************************************************************************************************************************
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeListadoCreditosReprogramados = sCadImp
End Function

'Creado por DAOR 20070418
Public Function nRepo108326_CreditosDesembConAprobExoneraReglamento(pdFecIni As Date, pdFecFin As Date, ByVal pMatAgencias As Variant, _
    ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "") As String
        
    Dim oDCred As DCredDoc
    Dim rs As ADODB.Recordset
    Dim sCadImp As String, sCadCabT As String
    Dim lsCadBuffer As String
    Dim lnIndice As Long
    Dim lnLineas As Integer, lnPage As Integer, I As Integer
    
             
    Set oDCred = New DCredDoc
    Set rs = oDCred.RecuperaCreditosDesembConAprobExoneraReglamento(pdFecIni, pdFecFin, pMatAgencias)
    Set oDCred = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        lnLineas = 0
        lnPage = 1

        'Cabecera
        sCadCab = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CREDITOS DESEMBOLSADOS CON APROBACIÓN DE EXONERACIÓN DE REGLAMENTO", psNomCmac, , , "108326", "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        sCadCabT = sCadCabT & String(120, "-") & lnSaltoLinDoc
        sCadCabT = sCadCabT & ImpreFormat("CREDITO", 18)
        sCadCabT = sCadCabT & ImpreFormat("CLIENTE", 35)
        sCadCabT = sCadCabT & ImpreFormat("FEC. DESEMB.", 10)
        sCadCabT = sCadCabT & ImpreFormat("MONEDA", 6)
        sCadCabT = sCadCabT & ImpreFormat("MONTO DESEMB.", 12)
        sCadCabT = sCadCabT & ImpreFormat("APODERADO", 10)
        sCadCabT = sCadCabT & ImpreFormat("AGENCIA", 15) & lnSaltoLinDoc
        sCadCabT = sCadCabT & String(120, "-") & lnSaltoLinDoc
        
        sCadCab = sCadImp + sCadCabT
        sCadImp = sCadImp + sCadCabT
        lnIndice = 0
        lnLineas = 7
        I = 0
        Do While Not rs.EOF
            I = I + 1
            lnIndice = lnIndice + 1
            lnLineas = lnLineas + 1
            
            'Cadena de Impresion
            sCadImp = sCadImp & ImpreFormat(rs!cCtaCod, 18)
            sCadImp = sCadImp & ImpreFormat(rs!cpersnombre, 35)
            sCadImp = sCadImp & ImpreFormat(Format(rs!dVigencia, "dd/mm/yyyy"), 10)
            sCadImp = sCadImp & ImpreFormat(IIf(rs!Moneda = 1, "Sóles", "Dólares"), 8)
            sCadImp = sCadImp & ImpreFormat(rs!nMontoCol, 10, 2, True)
            sCadImp = sCadImp & ImpreFormat(rs!cUser, 10)
            sCadImp = sCadImp & ImpreFormat(rs!cAgeDescripcion, 15)
            sCadImp = sCadImp & lnSaltoLinDoc
            
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & sCadImp
                sCadImp = ""
            End If
            
            rs.MoveNext
        Loop
        sCadImp = sCadImp & String(120, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & ImpreFormat("Total de Créditos : ", 20) & ImpreFormat(I, 9, 0, True) & lnSaltoLinDoc
       
        lsCadBuffer = lsCadBuffer & sCadImp & lnSaltoLinDoc
        nRepo108326_CreditosDesembConAprobExoneraReglamento = lsCadBuffer
    Else
        nRepo108326_CreditosDesembConAprobExoneraReglamento = ""
    End If
End Function

'Creado por DAOR 20070418
Public Function nRepo108327_MontoDesembolsadoPorLineas(pdFecIni As Date, pdFecFin As Date, ByVal pMatAgencias As Variant, _
    ByVal psCodUser As String, ByVal pdFecSis As Date, _
    ByVal psNomAge As String, Optional ByVal psNomCmac As String = "", Optional psCodRepo As String = "") As String
        
    Dim oDCred As DCredDoc
    Dim rs As ADODB.Recordset
    Dim sCadImp As String, sCadCabT As String
    Dim lsCadBuffer As String
    Dim lnIndice As Long
    Dim lnLineas As Integer, lnPage As Integer, I As Integer
    Dim nTotDesF1 As Double, nTotSalF1 As Double, nTotDesF2 As Double, nTotSalF2 As Double
    Dim nTotDesSF1 As Double, nTotSalSF1 As Double, nTotDesSF2 As Double, nTotSalSF2 As Double
    Dim nTotCasSF1 As Integer, nTotCasSF2 As Integer, nTotCasF1 As Integer, nTotCasF2 As Integer
    Dim sAgencia As String, sFondo As String, sSubFondo As String
             
    Set oDCred = New DCredDoc
    Set rs = oDCred.RecuperaMontosDesembolsadosPorLineas(pdFecIni, pdFecFin, pMatAgencias)
    Set oDCred = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        lnLineas = 0
        lnPage = 1

        'Cabecera
        sCadCab = ""
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "MONTOS DESEMBOLSADOS Y SALDO DE CAPITAL POR LINEAS", psNomCmac, , , "108327", "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        sCadCabT = sCadCabT & String(123, "-") & lnSaltoLinDoc
        sCadCabT = sCadCabT & Space(4) & ImpreFormat("LINEA", 13, 0)
        sCadCabT = sCadCabT & ImpreFormat("DESC. DE LINEA", 35)
        sCadCabT = sCadCabT & ImpreFormat("CASOS", 6)
        sCadCabT = sCadCabT & ImpreFormat("DESEMBOLSO", 12)
        sCadCabT = sCadCabT & ImpreFormat("SALDO", 12)
        sCadCabT = sCadCabT & ImpreFormat("CASOS", 6)
        sCadCabT = sCadCabT & ImpreFormat("DESEMBOLSO", 12)
        sCadCabT = sCadCabT & ImpreFormat("SALDO", 12) & lnSaltoLinDoc
        
        sCadCabT = sCadCabT & Space(56) & ImpreFormat("  MN  ", 6)
        sCadCabT = sCadCabT & ImpreFormat("     MN     ", 12)
        sCadCabT = sCadCabT & ImpreFormat("     MN     ", 12)
        sCadCabT = sCadCabT & ImpreFormat("  ME  ", 6)
        sCadCabT = sCadCabT & ImpreFormat("     ME     ", 12)
        sCadCabT = sCadCabT & ImpreFormat("     ME     ", 12) & lnSaltoLinDoc
        sCadCabT = sCadCabT & String(123, "-") & lnSaltoLinDoc
        
        sCadCab = sCadImp + sCadCabT
        sCadImp = sCadImp + sCadCabT
        lnIndice = 0
        lnLineas = 7
        I = 0
        sAgencia = rs!Agencia
        sFondo = rs!Fondo
        sSubFondo = rs!SubFondo
        
        sCadImp = sCadImp & rs!Agencia & lnSaltoLinDoc
        sCadImp = sCadImp & Space(2) & "Fondo : " & rs!Fondo & lnSaltoLinDoc
        sCadImp = sCadImp & Space(4) & "Sub Fondo : " & rs!SubFondo & lnSaltoLinDoc
        
        Do While Not rs.EOF
            I = I + 1
            lnIndice = lnIndice + 1
            lnLineas = lnLineas + 1
            If rs!SubFondo <> sSubFondo Then 'Mostrar resumen de SubFondo
                sCadImp = sCadImp & Space(4) & String(119, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & Space(4) & ImpreFormat("Total Sub Fondo " & sSubFondo & ": ", 49, 0)
                sCadImp = sCadImp & Space(1) & ImpreFormat(nTotCasSF1, 6, 0)
                sCadImp = sCadImp & Space(1) & ImpreFormat(nTotDesSF1, 10, 2, True)
                sCadImp = sCadImp & Space(1) & ImpreFormat(nTotSalSF1, 10, 2, True)
                sCadImp = sCadImp & Space(1) & ImpreFormat(nTotCasSF2, 6, 0)
                sCadImp = sCadImp & Space(1) & ImpreFormat(nTotDesSF2, 10, 2, True)
                sCadImp = sCadImp & Space(1) & ImpreFormat(nTotSalSF2, 10, 2, True) & lnSaltoLinDoc
                nTotCasSF1 = 0: nTotCasSF2 = 0
                nTotDesSF1 = 0: nTotDesSF2 = 0
                nTotSalSF1 = 0: nTotSalSF2 = 0
                sSubFondo = rs!SubFondo
                sCadImp = sCadImp & lnSaltoLinDoc
                If rs!Fondo = sFondo Then
                    sCadImp = sCadImp & Space(4) & "Sub Fondo : " & rs!SubFondo & lnSaltoLinDoc
                End If
            End If
                        
            If rs!Fondo <> sFondo Then 'Mostrar resumen de Fondo
                sCadImp = sCadImp & Space(2) & String(121, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & Space(2) & ImpreFormat("Total Fondo " & sFondo & ": ", 51, 0)
                sCadImp = sCadImp & Space(1) & ImpreFormat(nTotCasF1, 6, 0)
                sCadImp = sCadImp & Space(1) & ImpreFormat(nTotDesF1, 10, 2, True)
                sCadImp = sCadImp & Space(1) & ImpreFormat(nTotSalF1, 10, 2, True)
                sCadImp = sCadImp & Space(1) & ImpreFormat(nTotCasF2, 6, 0)
                sCadImp = sCadImp & Space(1) & ImpreFormat(nTotDesF2, 10, 2, True)
                sCadImp = sCadImp & Space(1) & ImpreFormat(nTotSalF2, 10, 2, True) & lnSaltoLinDoc
                nTotCasF1 = 0: nTotCasF2 = 0
                nTotDesF1 = 0: nTotDesF2 = 0
                nTotSalF1 = 0: nTotSalF2 = 0
                sFondo = rs!Fondo
                sCadImp = sCadImp & lnSaltoLinDoc
                If rs!Agencia = sAgencia Then
                    sCadImp = sCadImp & Space(2) & "Fondo : " & rs!Fondo & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(4) & "Sub Fondo : " & rs!SubFondo & lnSaltoLinDoc
                End If
            End If
            
            If rs!Agencia <> sAgencia Then 'Mostrar Resumen por Agencia
                sCadImp = sCadImp & String(123, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & lnSaltoLinDoc & lnSaltoLinDoc
                sCadImp = sCadImp & rs!Agencia & lnSaltoLinDoc
                sCadImp = sCadImp & Space(2) & "Fondo : " & rs!Fondo & lnSaltoLinDoc
                sCadImp = sCadImp & Space(4) & "Sub Fondo : " & rs!SubFondo & lnSaltoLinDoc
                sAgencia = rs!Agencia
            End If
            
            'Cadena de Impresion
            sCadImp = sCadImp & Space(4) & ImpreFormat(rs!cLineaCred, 13, 0)
            sCadImp = sCadImp & Space(1) & ImpreFormat(rs!Linea, 35, 0)
            sCadImp = sCadImp & Space(1) & ImpreFormat(rs!CasosMN, 6, 0)
            sCadImp = sCadImp & Space(1) & ImpreFormat(rs!DesembolsoMN, 10, 2, True)
            sCadImp = sCadImp & Space(1) & ImpreFormat(rs!SaldoMN, 10, 2, True)
            sCadImp = sCadImp & Space(1) & ImpreFormat(rs!CasosME, 6, 0)
            sCadImp = sCadImp & Space(1) & ImpreFormat(rs!DesembolsoME, 10, 2, True)
            sCadImp = sCadImp & Space(1) & ImpreFormat(rs!SaldoME, 10, 2, True) & lnSaltoLinDoc
            
            nTotCasSF1 = nTotCasSF1 + rs!CasosMN: nTotCasSF2 = nTotCasSF2 + rs!CasosME
            nTotDesSF1 = nTotDesSF1 + rs!DesembolsoMN: nTotDesSF2 = nTotDesSF2 + rs!DesembolsoME
            nTotSalSF1 = nTotSalSF1 + rs!SaldoMN: nTotSalSF2 = nTotSalSF2 + rs!SaldoME
            
            nTotCasF1 = nTotCasF1 + rs!CasosMN: nTotCasF2 = nTotCasF2 + rs!CasosME
            nTotDesF1 = nTotDesF1 + rs!DesembolsoMN: nTotDesF2 = nTotDesF2 + rs!DesembolsoME
            nTotSalF1 = nTotSalF1 + rs!SaldoMN: nTotSalF2 = nTotSalF2 + rs!SaldoME
            
            If lnIndice Mod 300 = 0 Then
                lsCadBuffer = lsCadBuffer & sCadImp
                sCadImp = ""
            End If
            
            rs.MoveNext
        Loop
        sCadImp = sCadImp & Space(4) & String(119, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(4) & ImpreFormat("Total Sub Fondo " & sSubFondo & ": ", 49, 0)
        sCadImp = sCadImp & Space(1) & ImpreFormat(nTotCasSF1, 6, 0)
        sCadImp = sCadImp & Space(1) & ImpreFormat(nTotDesSF1, 10, 2, True)
        sCadImp = sCadImp & Space(1) & ImpreFormat(nTotSalSF1, 10, 2, True)
        sCadImp = sCadImp & Space(1) & ImpreFormat(nTotCasSF2, 6, 0)
        sCadImp = sCadImp & Space(1) & ImpreFormat(nTotDesSF2, 10, 2, True)
        sCadImp = sCadImp & Space(1) & ImpreFormat(nTotSalSF2, 10, 2, True) & lnSaltoLinDoc
       
        sCadImp = sCadImp & Space(2) & String(121, "-") & lnSaltoLinDoc
        sCadImp = sCadImp & Space(2) & ImpreFormat("Total Fondo " & sFondo & ": ", 51, 0)
        sCadImp = sCadImp & Space(1) & ImpreFormat(nTotCasF1, 6, 0)
        sCadImp = sCadImp & Space(1) & ImpreFormat(nTotDesF1, 10, 2, True)
        sCadImp = sCadImp & Space(1) & ImpreFormat(nTotSalF1, 10, 2, True)
        sCadImp = sCadImp & Space(1) & ImpreFormat(nTotCasF2, 6, 0)
        sCadImp = sCadImp & Space(1) & ImpreFormat(nTotDesF2, 10, 2, True)
        sCadImp = sCadImp & Space(1) & ImpreFormat(nTotSalF2, 10, 2, True) & lnSaltoLinDoc
       
        lsCadBuffer = lsCadBuffer & sCadImp & lnSaltoLinDoc
        nRepo108327_MontoDesembolsadoPorLineas = lsCadBuffer
    Else
        nRepo108327_MontoDesembolsadoPorLineas = ""
    End If
End Function

'**DAOR 20070717
Public Function ImprimeConsolidadoCarteraXAnalista(ByVal pMatAgencias As Variant, ByVal pdFecIni As Date, ByVal pdFecFin As Date, _
    ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal psNomAge As String, ByVal pMatProd As Variant, _
    Optional ByVal psNomCmac As String = "", Optional ByVal pnTipCam As Double, Optional ByVal psTitProductos As String, Optional ByVal pbCondBN As Boolean = False) As String
Dim oDCred As DCredDoc
Dim R As ADODB.Recordset
Dim sCadImp As String

    Set oDCred = New DCredDoc
    Set R = oDCred.RecuperaConsolidadoCarteraPorAnalista(pMatAgencias, pdFecIni, pdFecFin, pMatProd, pnTipCam)
    Set oDCred = Nothing
    
    sCadImp = ""
    If R.RecordCount > 0 Then
        Call ImprimeCabeceraDocumento(sCadImp, psNomAge, pdFecSis, psCodUser, "CONSOLIDADO DE CARTERA POR ANALISTA", psNomCmac, , True, , "DEL " & Format(pdFecIni, "dd/mm/yyyy") & " AL " & Format(pdFecFin, "dd/mm/yyyy"))
        sCadImp = sCadImp & "Agencia     : " & R!cAgencia & lnSaltoLinDoc
        sCadImp = sCadImp & "Tipo Cambio : " & Str(pnTipCam) & lnSaltoLinDoc
        sCadImp = sCadImp & "Productos   : " & psTitProductos
        sCadImp = sCadImp & lnSaltoLinDoc
        sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
        
        sCadImp = sCadImp & ImpreFormat("ANAL", 7, 0)
        sCadImp = sCadImp & ImpreFormat("  SALDO CARTERA", 17) & "|"
        sCadImp = sCadImp & ImpreFormat("  NUEVOS", 17) & "|"
        sCadImp = sCadImp & ImpreFormat("  RECURRENTES", 17) & "|"
        sCadImp = sCadImp & ImpreFormat("  PARALELOS", 17) & "|"
        sCadImp = sCadImp & ImpreFormat("  REFINANCIADOS", 17) & "|"
        sCadImp = sCadImp & ImpreFormat("  AMPLIADO", 17) & "|"
        sCadImp = sCadImp & ImpreFormat("  AUTOMATICOS", 17) & "|" & lnSaltoLinDoc
        sCadImp = sCadImp & Space(7)
        sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|"
        sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|" & lnSaltoLinDoc
        sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
    End If
    Do While Not R.EOF
            'FINAL DE AGENCIA
            If R!cUser = "TOTAL" Then
                sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
                sCadImp = sCadImp & ImpreFormat("SUB T:", 6, 0) '
                sCadImp = sCadImp & ImpreFormat(R!nCantNuevo + R!nCantRecurrente + R!nCantParalelo + R!nCantRefinanciado + R!nCantAmpliado + R!nCantAutomatico, 6, 0) & Space(1) & ImpreFormat(R!nMontoNuevo + R!nMontoRecurrente + R!nMontoParalelo + R!nMontoRefinanciado + R!nMontoAmpliado + R!nMontoAutomatico, 10, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCantNuevo, 6, 0) & Space(1) & ImpreFormat(R!nMontoNuevo, 10, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCantRecurrente, 6, 0) & Space(1) & ImpreFormat(R!nMontoRecurrente, 10, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCantParalelo, 6, 0) & Space(1) & ImpreFormat(R!nMontoParalelo, 10, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCantRefinanciado, 6, 0) & Space(1) & ImpreFormat(R!nMontoRefinanciado, 10, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCantAmpliado, 6, 0) & Space(1) & ImpreFormat(R!nMontoAmpliado, 10, 2)
                sCadImp = sCadImp & ImpreFormat(R!nCantAutomatico, 6, 0) & Space(1) & ImpreFormat(R!nMontoAutomatico, 10, 2)
                sCadImp = sCadImp & lnSaltoLinDoc
                sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
                R.MoveNext
                If R.EOF Then
                    Exit Do
                End If
                'FINAL FINAL
                If R!cUser = "TOTAL" Then
                    sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("TOTAL:", 6, 0)
                    sCadImp = sCadImp & ImpreFormat(R!nCantNuevo + R!nCantRecurrente + R!nCantParalelo + R!nCantRefinanciado + R!nCantAmpliado + R!nCantAutomatico, 6, 0) & Space(1) & ImpreFormat(R!nMontoNuevo + R!nMontoRecurrente + R!nMontoParalelo + R!nMontoRefinanciado + R!nMontoAmpliado + R!nMontoAutomatico, 10, 2)
                    sCadImp = sCadImp & ImpreFormat(R!nCantNuevo, 6, 0) & Space(1) & ImpreFormat(R!nMontoNuevo, 10, 2)
                    sCadImp = sCadImp & ImpreFormat(R!nCantRecurrente, 6, 0) & Space(1) & ImpreFormat(R!nMontoRecurrente, 10, 2)
                    sCadImp = sCadImp & ImpreFormat(R!nCantParalelo, 6, 0) & Space(1) & ImpreFormat(R!nMontoParalelo, 10, 2)
                    sCadImp = sCadImp & ImpreFormat(R!nCantRefinanciado, 6, 0) & Space(1) & ImpreFormat(R!nMontoRefinanciado, 10, 2)
                    sCadImp = sCadImp & ImpreFormat(R!nCantAmpliado, 6, 0) & Space(1) & ImpreFormat(R!nMontoAmpliado, 10, 2)
                    sCadImp = sCadImp & ImpreFormat(R!nCantAutomatico, 6, 0) & Space(1) & ImpreFormat(R!nMontoAutomatico, 10, 2)
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
                    Exit Do
                Else
                    sCadImp = sCadImp & Chr$(12)
                    sCadImp = sCadImp & "Agencia     : " & R!cAgencia & lnSaltoLinDoc
                    sCadImp = sCadImp & "Tipo Cambio : " & Str(pnTipCam) & lnSaltoLinDoc
                    sCadImp = sCadImp & "Productos   : " & psTitProductos
                    sCadImp = sCadImp & lnSaltoLinDoc
                    sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
                    sCadImp = sCadImp & ImpreFormat("ANAL", 7, 0)
                    sCadImp = sCadImp & ImpreFormat("  SALDO CARTERA", 17) & "|"
                    sCadImp = sCadImp & ImpreFormat("  NUEVOS", 17) & "|"
                    sCadImp = sCadImp & ImpreFormat("  RECURRENTES", 17) & "|"
                    sCadImp = sCadImp & ImpreFormat("  PARALELOS", 17) & "|"
                    sCadImp = sCadImp & ImpreFormat("  REFINANCIADOS", 17) & "|"
                    sCadImp = sCadImp & ImpreFormat("  AMPLIADO", 17) & "|"
                    sCadImp = sCadImp & ImpreFormat("  AUTOMATICOS", 17) & "|" & lnSaltoLinDoc
                    sCadImp = sCadImp & Space(7)
                    sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|"
                    sCadImp = sCadImp & ImpreFormat("NUM.   MONTO     ", 17) & "|" & lnSaltoLinDoc
                    sCadImp = sCadImp & String(150, "-") & lnSaltoLinDoc
                End If
            End If
            sCadImp = sCadImp & ImpreFormat(R!cUser, 6, 0)
            sCadImp = sCadImp & ImpreFormat(R!nCantNuevo + R!nCantRecurrente + R!nCantParalelo + R!nCantRefinanciado + R!nCantAmpliado + R!nCantAutomatico, 6, 0) & Space(1) & ImpreFormat(R!nMontoNuevo + R!nMontoRecurrente + R!nMontoParalelo + R!nMontoRefinanciado + R!nMontoAmpliado + R!nMontoAutomatico, 10, 2)
            sCadImp = sCadImp & ImpreFormat(R!nCantNuevo, 6, 0) & Space(1) & ImpreFormat(R!nMontoNuevo, 10, 2)
            sCadImp = sCadImp & ImpreFormat(R!nCantRecurrente, 6, 0) & Space(1) & ImpreFormat(R!nMontoRecurrente, 10, 2)
            sCadImp = sCadImp & ImpreFormat(R!nCantParalelo, 6, 0) & Space(1) & ImpreFormat(R!nMontoParalelo, 10, 2)
            sCadImp = sCadImp & ImpreFormat(R!nCantRefinanciado, 6, 0) & Space(1) & ImpreFormat(R!nMontoRefinanciado, 10, 2)
            sCadImp = sCadImp & ImpreFormat(R!nCantAmpliado, 6, 0) & Space(1) & ImpreFormat(R!nMontoAmpliado, 10, 2)
            sCadImp = sCadImp & ImpreFormat(R!nCantAutomatico, 6, 0) & Space(1) & ImpreFormat(R!nMontoAutomatico, 10, 2)
            sCadImp = sCadImp & lnSaltoLinDoc
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    ImprimeConsolidadoCarteraXAnalista = sCadImp
    
End Function
