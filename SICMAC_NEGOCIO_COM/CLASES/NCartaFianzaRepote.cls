VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCartaFianzaReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim csNomCMAC As String
Dim csNomAgencia As String
Dim csCodUser As String
Dim csFechaSis As String
Public Function nRepoTarifario() As String
Dim Rs As New ADODB.Recordset
Dim CF As DCartaFianza
Dim lsCadena As String
Dim Tabula As String
Dim x As String
Set CF = New DCartaFianza
Set Rs = CF.RecuperaCF_Tarifario
Tabula = Space(5)
lsCadena = String(80, "-") & Chr(10)
lsCadena = lsCadena & Tabula & "Cod. Tarif." & Space(2) & "Modalidad" & Space(20) & "Moneda" & Space(5) & "Tasa Trim." & Space(2) & "Monto Minimo" & Chr(10)
lsCadena = lsCadena & String(80, "-") & Chr(10)
While Not Rs.EOF
    lsCadena = lsCadena & Tabula & ImpreFormat(Rs!ctarifCod, 5) & Space(3) & ImpreFormat(Rs!Modalidad, 25) & Space(5) & ImpreFormat(Rs!Moneda, 7) & Space(3) & ImpreFormat(Rs!nTasaTrim, 4, 2, False) & Space(4) & ImpreFormat(Rs!nMontoMinimo, 8, 2, False) & Chr(10)
    Rs.MoveNext
Wend
nRepoTarifario = lsCadena
Set Rs = Nothing
Set CF = Nothing
End Function


Public Function nImprimeCartaFianza(ByVal pnCodCta As String) As String
Dim Rs As ADODB.Recordset
Dim oCon As DConecta
Dim sql  As String
Set oCon = New DConecta
Set Rs = New ADODB.Recordset

oCon.AbreConexion
sql = "SELECT GETDATE() "
Set Rs = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing

End Function
Public Function nRepoDuplicado(ByVal pnCodCta As String, ByRef SoSu As Integer) As String
Dim loRs As NCartaFianzaValida

Dim lrDataT As ADODB.Recordset
Dim lrDataCR As ADODB.Recordset
Dim lrDataCF As ADODB.Recordset

Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lnNegritaON As String
Dim lnNegritaOFF As String
    
    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr(27) & Chr$(70)
    
    Set loRs = New NCartaFianzaValida
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(pnCodCta)
    Set lrDataT = loRs.RecuperaDatosT(pnCodCta)
    
    If lrDataT Is Nothing Or (lrDataT.BOF And lrDataT.EOF) Then
        MsgBox " No Existen Datos para el reporte", vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1
        'CABECERA
        Select Case SoSu
            Case 2
            lsCadImp = lsCadImp & nRepoCabecera("R E G I S T R O   D E   S O L I C I T U D   D E    C A R T A   F I A N Z A", " ", lnPage, 130, "", 100)
            Case 3
            lsCadImp = lsCadImp & nRepoCabecera("R E G I S T R O   D E   S U G E R E N C I A   D E   A N A L I S T A   P A R A   C A R T A   F I A N Z A", " ", lnPage, 130, "", 100)
        End Select
        
        'lsCadImp = lsCadImp & nRepoCabecera("R E S U M E N   D E   C A R T A    F I A N Z A ", " ", lnPage, 130, "", 100)
        lnIndice = 0:  lnLineas = 7
        '***********************************************
        
        lsCadImp = lsCadImp & Space(20) & "CARTA FIANZA:" & Space(5) & lrDataCF!Carta
        lsCadImp = lsCadImp & Space(30) & lrDataCF!Tipo & Chr(10)
        
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(10) & ImpreFormat("  DATOS DEL TITULAR", 40, 0) & Chr(10)
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lnLineas = lnLineas + 4
        '***************** DATOS TITULAR
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("CLIENTE", 20, 0) & ":" & Space(2) & lrDataT!Nombre & Chr(10)
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("DOC. IDENTIDAD", 20, 0) & ":" & Space(2) & lrDataT!DocDNI & Chr(10)
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("DOC. TRIBUTARIO", 20, 0) & ":" & Space(2) & lrDataT!DocRUC & Chr(10)
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("DIRECCION", 20, 0) & ":" & Space(2) & lrDataT!Direccion & Chr(10)
        lnLineas = lnLineas + 4
        
        If SoSu = 3 Then
         Set lrDataCR = loRs.RecuperaDatosJuridica(lrDataT!Codigo)
         'MsgBox lrDataT!Codigo

         If lrDataCR.RecordCount = 0 Or (lrDataCR.EOF And lrDataCR.BOF) Then
         Else
         lsCadImp = lsCadImp & Space(2) & ImpreFormat("TIEMPO CREACION", 20, 0) & ":" & Space(2) & lrDataT!Nacimiento & Chr(10)
         lsCadImp = lsCadImp & Space(2) & ImpreFormat("TIPO JURIDICO", 20, 0) & ":" & Space(2) & "" & lrDataCR!DescribeJur & Chr(10)
         lsCadImp = lsCadImp & Space(2) & ImpreFormat("MAGNITUD EMPRESA", 20, 0) & ":" & Space(2) & "" & lrDataCR!Magnitud & Chr(10)
         lnLineas = lnLineas + 3
         End If
        End If
        
        '--------------------------------------------------
        ' DATOS ACREEDOR
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(10) & ImpreFormat("  DATOS DEL ACREEDOR", 40, 0) & Chr(10)
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lnLineas = lnLineas + 3
        
        Set lrDataCR = loRs.RecuperaDatosAcreedor(pnCodCta)
        If lrDataCR.EOF And lrDataCR.BOF Then
        Else
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("CLIENTE", 20, 0) & ":" & Space(2) & lrDataCR!Nombre & Chr(10)
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("DOC. IDENTIDAD", 20, 0) & ":" & Space(2) & lrDataCR!DocDNI & Chr(10)
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("DOC. TRIBUTARIO", 20, 0) & ":" & Space(2) & lrDataT!DocRUC & Chr(10)
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("DIRECCION", 20, 0) & ":" & Space(2) & lrDataCR!Direccion & Chr(10)
        lnLineas = lnLineas + 4
        End If
        '-------------------------------------------------
        
        'DATOS FUENTE DE INGRESO
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(10) & ImpreFormat("  FUENTE DE INGRESO", 40, 0) & Chr(10)
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lnLineas = lnLineas + 3
        
        Set lrDataCR = loRs.RecuperaFuenteIngreso(lrDataT!Codigo)
        If lrDataCR.EOF And lrDataCR.BOF Then
        Else
         lsCadImp = lsCadImp & Space(2) & ImpreFormat("Tipo Ingreso", 20, 0) & ":" & Space(2) & lrDataCR!Tipo_Ingreso & Chr(10)
         lsCadImp = lsCadImp & Space(2) & ImpreFormat("Nombre/Razon Social", 20, 0) & ":" & Space(2) & lrDataCR!Razon & Chr(10)
         lsCadImp = lsCadImp & Space(2) & ImpreFormat("Direccion", 20, 0) & ":" & Space(2) & lrDataCR!Direccion & Chr(10)
         lsCadImp = lsCadImp & Space(2) & ImpreFormat("Actividad", 20, 0) & ":" & Space(2) & lrDataCR!Actividad & Chr(10)
         lsCadImp = lsCadImp & Space(2) & ImpreFormat("Ultima Actualizacion", 20, 0) & ":" & Space(2) & Format(lrDataCR!UltimaAct, "dd/mm/yyyy") & Chr(10)
         
         lnLineas = lnLineas + 5
        End If
        
        '--------------------------------------------------
        'DATOS CLIENTES RELACIONADOS
        'lsCadImp = lsCadImp & String(130, "-") & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Space(10) & ImpreFormat("  CLIENTES RELACIONADOS", 40, 0) & Chr(10)
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lnLineas = lnLineas + 2
        
        lsCadImp = lsCadImp & Space(4) & "CODIGO" & Space(10) & "NOMBRE" & Space(42) & "RELACIONES" & Chr(10)
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lnLineas = lnLineas + 2
        Set lrDataCR = loRs.RecuperaPersonaRelacion(pnCodCta)
        
        While Not lrDataCR.EOF
            lsCadImp = lsCadImp & Space(4) & lrDataCR!Codpersona
            lsCadImp = lsCadImp & Space(2) & ImpreFormat(lrDataCR!Nombre, 40, 0) & Space(10)
            lsCadImp = lsCadImp & Space(2) & ImpreFormat(lrDataCR!Relacion, 30, 0) & Chr(10)
            lrDataCR.MoveNext
            lnLineas = lnLineas + 1
         Wend
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lnLineas = lnLineas + 1
        '----------------------------------------------------
        'DATOS CARTA FIANZA
        
        'lsCadImp = lsCadImp & String(130, "-") & Chr(10) & Chr(10)
        lsCadImp = lsCadImp & Space(10) & ImpreFormat("  DATOS DE LA CARTA FIANZA", 40, 0) & Chr(10)
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lnLineas = lnLineas + 2
        
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("Analista", 20, 0) & ":" & Space(2) & ImpreFormat(lrDataCF!Analista, 30, 0) & Space(10)
        lsCadImp = lsCadImp & ImpreFormat("Fuente de Ingreso", 20, 0) & ":" & Space(2) & lrDataCF!F_Ingreso & Chr(10)
        
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("Fecha Asignacion", 20, 0) & ":" & Space(2) & ImpreFormat(Format(lrDataCF!F_Asignacion, "dd/mm/yyyy"), 30, 0) & Space(10)
        lsCadImp = lsCadImp & ImpreFormat("Modalidad", 20, 0) & ":" & Space(2) & lrDataCF!Modalidad & Chr(10)
        
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("Fecha de Vigencia", 20, 0) & ":" & Space(2) & ImpreFormat(Format(lrDataCF!Vence, "dd/mm/yyyy"), 30, 0) & Space(10)
        lsCadImp = lsCadImp & ImpreFormat("Condicion", 20, 0) & ":" & Space(2) & lrDataCF!Condicion & Chr(10)
        
        If SoSu = 2 Then
            lsCadImp = lsCadImp & Space(2) & ImpreFormat("Monto Solicitado", 20, 0) & ":" & Space(2) & IIf(Mid(lrDataCF!Carta, 9, 1) = "1", "S/. ", "$ ") & Format(lrDataCF!Monto, "##,##0.00") & Chr(10)
        End If
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lnLineas = lnLineas + 5
        '----------------------------------------------------------------
        'RELAC INTERNAAS
        
        If SoSu = 3 Then
        'lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lsCadImp = lsCadImp & Space(10) & ImpreFormat("  RELACIONES INTERNAS", 40, 0) & Chr(10)
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lnLineas = lnLineas + 2
        Set lrDataCR = loRs.RecuperaRelacinesInternas(lrDataT!Codigo)
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("CLIENTE ", 30, 0) & Space(2)
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("RELACION ", 15, 0) & Space(4)
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("CARTA FIANZA ", 20, 0) & Space(2)
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("MONTO ", 10, 0) & Chr(10)
        lsCadImp = lsCadImp & String(130, "-") & Chr(10)
        lnLineas = lnLineas + 2
        While Not lrDataCR.EOF
            lsCadImp = lsCadImp & Space(2) & ImpreFormat(lrDataCR!Cliente, 30, 0) & Space(2)
            lsCadImp = lsCadImp & ImpreFormat(lrDataCR!Relacion, 15, 0) & Space(2)
            lsCadImp = lsCadImp & ImpreFormat(lrDataCR!Carta, 20, 0) & Space(2)
            lsCadImp = lsCadImp & ImpreFormat(lrDataCR!Saldo, 10, 2) & Space(1) & lrDataCR!Moneda & Chr(10)
            lrDataCR.MoveNext
            lnLineas = lnLineas + 1
         Wend
        
        End If


        If lnLineas >= 55 Then
           lnPage = lnPage + 1
           lsCadImp = lsCadImp & Chr(12)
           Select Case SoSu
            Case 2
                lsCadImp = lsCadImp & nRepoCabecera("R E G I S T R O   D E   S O L I C I T U D   D E    C A R T A   F I A N Z A", "  ", lnPage, 130, "", 100)
            Case 3
                lsCadImp = lsCadImp & nRepoCabecera("R E G I S T R O   D E   S U G E R E N C I A   D E   A N A L I S T A   P A R A   C A R T A   F I A N Z A", " ", lnPage, 130, "", 100)
            End Select
           lnLineas = 7
        End If
       
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    End If
    nRepoDuplicado = lsCadBuffer


End Function

Public Function nRepoCartaFianza_Estados(ByVal pnCodCta As String) As String

Dim lrDataRep As ADODB.Recordset
Dim lrDataT As ADODB.Recordset
Dim lrDataH As ADODB.Recordset
Dim lrDataG As ADODB.Recordset
Dim lrDataO As ADODB.Recordset

'Dim loDataRep As dColPFunciones
Dim loRs As NCartaFianzaValida

Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lnNegritaON As String
Dim lnNegritaOFF As String
    
    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr(27) & Chr$(70)
    
    
    Set loRs = New NCartaFianzaValida
    Set lrDataRep = loRs.RecuperaDatosGeneralesCF(pnCodCta)
    Set lrDataT = loRs.RecuperaPersonaRelacion(pnCodCta)
    
    'If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
    '    MsgBox " No Existen Datos para el reporte" & psAgencia, vbInformation, " Aviso"
    '    Exit Function
    'Else
        lnLineas = 0
        lnPage = 1
        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("R E S U M E N   D E   C A R T A    F I A N Z A ", " ", lnPage, 150, "", 100)
        lnIndice = 0:  lnLineas = 7
        lsCadImp = lsCadImp & lnNegritaON
        lsCadImp = lsCadImp & "Nª Carta Fianza : " & lrDataRep!Carta & Space(15) & "Tipo : " & lrDataRep!Tipo & _
                    Space(15) & "Estado Actual : " & lrDataRep!Estado & Chr(10)
        lsCadImp = lsCadImp & lnNegritaOFF
        '***********************************************
        
        lsCadImp = lsCadImp & String(150, "-") & Chr(10) & Chr(10)
        
        lsCadImp = lsCadImp & Space(10) & ImpreFormat("  DATOS GENERALES", 40, 0) & Chr(10)
        lsCadImp = lsCadImp & String(150, "-") & Chr(10)
        
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("Analista", 20, 0) & ":" & Space(2) & ImpreFormat(lrDataRep!Analista, 30, 0) & Space(10)
        lsCadImp = lsCadImp & ImpreFormat("Fuente de Ingreso", 20, 0) & ":" & Space(2) & lrDataRep!F_Ingreso & Chr(10)
        
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("Apoderado", 20, 0) & ":" & Space(2) & ImpreFormat(IIf(IsNull(lrDataRep!Apoderado), "", lrDataRep!Apoderado), 30, 0) & Space(10)
        lsCadImp = lsCadImp & ImpreFormat("Fecha de Vigencia", 20, 0) & ":" & Space(2) & Format(lrDataRep!Vence, "dd/mm/yyyy") & Chr(10)
        
        lsCadImp = lsCadImp & Space(2) & ImpreFormat("Modalidad", 20, 0) & ":" & Space(2) & ImpreFormat(lrDataRep!Modalidad, 30, 0) & Space(10)
        lsCadImp = lsCadImp & ImpreFormat("Condicion", 20, 0) & ":" & Space(2) & lrDataRep!Condicion & Chr(10)
        
        lsCadImp = lsCadImp & String(150, "-") & Chr(10)
        lnLineas = lnLineas + 7
        
        Set lrDataT = loRs.RecuperaPersonaRelacion(pnCodCta)
         lsCadImp = lsCadImp & Space(10) & "NOMBRE" & Space(37) & "RELACIONES" & Chr(10)
         lsCadImp = lsCadImp & String(150, "-") & Chr(10)
         lnLineas = lnLineas + 2
         
         While Not lrDataT.EOF
            lsCadImp = lsCadImp & Space(2) & ImpreFormat(lrDataT!Nombre, 40, 0) & Space(10)
            lsCadImp = lsCadImp & Space(2) & ImpreFormat(lrDataT!Relacion, 30, 0) & Chr(10)
            lrDataT.MoveNext
            lnLineas = lnLineas + 1
         Wend
        lsCadImp = lsCadImp & String(150, "-") & Chr(10) & Chr(10)
        
        '-------------------------------
        Set lrDataRep = loRs.RecuperaHistorial(pnCodCta)
        Set lrDataH = loRs.RecuperaHistorial(pnCodCta)
        
        lsCadImp = lsCadImp & Space(53) & "FECHA" & Space(28) & "MONTO" & Chr(10)
        lsCadImp = lsCadImp & String(150, "-") & Chr(10)
         lnLineas = lnLineas + 3
         While Not lrDataH.EOF
            lsCadImp = lsCadImp & Space(2) & ImpreFormat(lrDataH!Estado, 10, 0)
            lsCadImp = lsCadImp & Space(42) & Format(lrDataH!Fecha, "dd/mm/yyyy")
            lsCadImp = lsCadImp & Space(10) & ImpreFormat(lrDataH!Monto, 15, 2) & Chr(10)
            lrDataH.MoveNext
            lnLineas = lnLineas + 1
         Wend
         
         lsCadImp = lsCadImp & String(150, "-") & Chr(10) & Chr(10)
         lnLineas = lnLineas + 2
        
         '-------------------------------
         
         Set lrDataG = loRs.RecuperaGarantias(pnCodCta)
         
        lsCadImp = lsCadImp & Space(2) & "Nª Garantia" & Space(12)
        lsCadImp = lsCadImp & "Tipo Garantia" & Space(15) & "Descripcion" & Space(21) & "Documento" & Chr(10)
        lsCadImp = lsCadImp & String(150, "-") & Chr(10)
        
        
        lnLineas = lnLineas + 3
        While Not lrDataG.EOF
            lsCadImp = lsCadImp & Space(2) & ImpreFormat(lrDataG!NroGarantia, 10, 0)
            lsCadImp = lsCadImp & Space(13) & Format(lrDataG!Tipo_Garn, "dd/mm/yyyy")
            lsCadImp = lsCadImp & Space(20) & ImpreFormat(lrDataG!Descripcion, 20, 0)
            lsCadImp = lsCadImp & Space(12) & ImpreFormat(lrDataG!Tipo_Doc, 15, 0) & Chr(10)
            lrDataG.MoveNext
            lnLineas = lnLineas + 1
        Wend
        lsCadImp = lsCadImp & String(150, "-") & Chr(10) & Chr(10)
        lnLineas = lnLineas + 2
        
        
        '-----------------------------------------------
        Set lrDataO = loRs.RepcuperaOtrosDt(pnCodCta)
        If lrDataO Is Nothing Or (lrDataO.EOF And lrDataO.BOF) Then
        Else
         lsCadImp = lsCadImp & String(150, "-") & Chr(10)
         lsCadImp = lsCadImp & lnNegritaON
         lsCadImp = lsCadImp & "Motivo Rechazo :" & ImpreFormat(lrDataO!Motivo, 20, 0) & Space(2)
         lsCadImp = lsCadImp & "Fecha :" & Format(lrDataO!Fecha, "dd/mm/yyyy") & Chr(10)
         lsCadImp = lsCadImp & lnNegritaOFF
         lsCadImp = lsCadImp & String(150, "-") & Chr(10)
         lnLineas = lnLineas + 3
        End If
        
        '----------------------------------------------------
        If lnLineas >= 55 Then
           lnPage = lnPage + 1
           lsCadImp = lsCadImp & Chr(12)
           lsCadImp = lsCadImp & nRepoCabecera("R E S U M E N   D E   C A R T A    F I A N Z A ", " ", lnPage, 150, "", 100)
           lnLineas = 7
        End If
       
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
    '-End if
    nRepoCartaFianza_Estados = lsCadBuffer


End Function
Public Sub inicio(ByVal psNomCmac As String, ByVal psNomAgencia As String, _
            ByVal psCodUser As String, ByVal psFechaSis As String)
    csNomCMAC = psNomCmac
    csNomAgencia = psNomAgencia
    csCodUser = psCodUser
    csFechaSis = psFechaSis
End Sub
Public Function nRepoCabecera(ByVal psTitulo As String, ByVal psSubTitulo As String, _
        ByVal pnPagina As Integer, ByVal pnAnchoLinea As Integer, ByVal psComenta As String, _
        ByVal pnCodReporte As Long) As String
        
Dim lsCadImp As String
Dim loImprimeCab As NColPImpre
    
    Set loImprimeCab = New NColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    lsCadImp = lsCadImp & Chr(10) & String(pnAnchoLinea, "-") & Chr(10)
    Select Case pnCodReporte
        Case 148001  ' Cartas Fianzas Aprobadas
            lsCadImp = lsCadImp & Space(1) & "Item   CartaFianza           Cliente                        Emision  Vencimiento    Acreedor                                Analista   Estado" & Chr(10)
        Case 148002 ' Cartas Fianzas Emitidas
            lsCadImp = lsCadImp & Space(1) & "Item   CartaFianza           Cliente                        Emision  Vencimiento    Acreedor                                Analista   " & Chr(10)
        Case 148003 ' Cartas Fianzas Vigentes
            lsCadImp = lsCadImp & Space(1) & "Item   CartaFianza           Cliente                        Emision  Vencimiento      Importe    Acreedor                           Analista " & Chr(10)
        Case 148004 ' Cartas Fianzas Vencidas
            lsCadImp = lsCadImp & Space(1) & "Item   CartaFianza           Cliente                        Emision  Vencimiento    Acreedor                                Analista   Carta Honrada" & Chr(10)
        Case 148005 ' Cartas Fianzas Devueltas
            lsCadImp = lsCadImp & Space(1) & "Item   CartaFianza           Cliente                        Emision  Vencimiento    Acreedor                                Analista   " & Chr(10)
    End Select
    lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & Chr(10)
nRepoCabecera = lsCadImp
End Function
Public Function nRepo148002_CartasFianzaVigentes(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdCMACT As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoC As Integer, ByRef psTipoM As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String) As String
        
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim P As String

lsSQL = "Select CF.cCtaCod as Carta_Fianza, CF.dAsignacion as Fecha_Emision, CF.dVencimiento as Fecha_Vencimiento," _
    & " P1.cPersNombre as Cliente, P2.cPersNombre as Acreedor, CC.nMontoCol as Importe," _
    & " RH.cUser as Analista, " _
    & " Ag.cAgeDescripcion as NombAgencia, " _
    & " ( Select C.cConsDescripcion  from Constante C where Pro.nPrdEstado=C.nConsValor and nConsCod=3001) as Estado " _
    & " From Producto Pro " _
    & " Inner Join colocaciones CC on CC.cCtaCod = Pro.cCtaCod " _
    & " " _
    & " Inner Join ColocCartaFianza CF on CF.cCtaCod = Pro.cCtaCod " _
    & " Inner Join ProductoPersona PP  on CF.cCtaCod=PP.cCtaCod " _
    & " Inner Join ProductoPersona PP2 on CF.cCtaCod=PP2.cCtaCod " _
    & " Inner Join ProductoPersona PP3 on CF.cCtaCod=PP3.cCtaCod " _
    & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod " _
    & " Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod " _
    & " Inner Join RRHH    RH on RH.cPersCod = PP3.cPersCod " _
    & " Inner Join Agencias Ag on Ag.cAgeCod=Substring(CF.cCtaCod,4,2)" _
    & " WHERE PP.nPrdPersRelac = 20 and PP2.nPrdPersRelac=35 and PP3.nPrdPersRelac=28 " _
    & " And Pro.nPrdEstado=" & gColocEstVigNorm & "" _
    & "  "

    If OptAge = True Then
        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)='" & psAgencia & "'"
    Else
        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2) in (" & psListaAgencias & " )"
    End If
    'lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)=" & psAgencia
                            
    
   '************
   If psTipoC = 0 And psTipoM = 1 Then 'tipo = Comercial
       lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 2 & "' "
   End If
    
   If psTipoC = 1 And psTipoM = 0 Then 'tipo = MicroEmpresa
      lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 1 & "' "
   End If
   '******************
   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
    lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 1 & "' "
   End If
   If psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
    lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 2 & "' "
   End If
        
'On Error GoTo dError

    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !Carta_Fianza & Space(1) & ImpreFormat(!Cliente, 30) & Space(2) _
                    & !Fecha_Emision & Space(2) & !Fecha_Vencimiento & Space(2) & ImpreFormat(!Acreedor, 40) _
                    & Space(2) & !Analista
                
                If Not OptAge Then
                    lsCadImp = lsCadImp & Space(2) & !NombAgencia
                End If
                lsCadImp = lsCadImp & Chr(10)

                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
       
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo148002_CartasFianzaVigentes = lsCadBuffer

End Function
Public Function nRepo148003_CartasFianzasVigentes_x_Vencer(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoC As Integer, ByRef psTipoM As Integer, ByRef pdCMACT As Date, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String) As String
        

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim P As String

lsSQL = "Select CF.cCtaCod as Carta_Fianza, CF.dAsignacion as Fecha_Emision, CF.dVencimiento as Fecha_Vencimiento," _
    & " P1.cPersNombre as Cliente, P2.cPersNombre as Acreedor, Pro.nSaldo as Importe," _
    & " RH.cUser as Analista, Pro.dPrdEstado, " _
    & " Ag.cAgeDescripcion as NombAgencia " _
    & " From ColocCartaFianza CF " _
    & " Inner Join ColocCFEstado CFE on CFE.cCtaCod = CF.cCtaCod " _
    & " Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod " _
    & " Inner Join ProductoPersona PP  on CF.cCtaCod=PP.cCtaCod " _
    & " Inner Join ProductoPersona PP2 on CF.cCtaCod=PP2.cCtaCod " _
    & " Inner Join ProductoPersona PP3 on CF.cCtaCod=PP3.cCtaCod " _
    & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod " _
    & " Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod " _
    & " Inner Join RRHH    RH on RH.cPersCod = PP3.cPersCod " _
    & " Inner Join Agencias Ag on Ag.cAgeCod=Substring(CF.cCtaCod,4,2)" _
    & " WHERE PP.nPrdPersRelac = 20 and PP2.nPrdPersRelac=35 and PP3.nPrdPersRelac=28 and CFE.nPrdEstado=2020 " _
    & "  " _
    & " and CF.dVencimiento between convert(datetime,'" & Format(pdDesde, "mm/dd/yyyy") & "') and  convert(datetime,'" & Format(pdHasta, "mm/dd/yyyy") & "')"
    
    If OptAge = True Then
        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)= '" & psAgencia & "'"
    Else
        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2) in (" & psListaAgencias & " )"
    End If
    'lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)=" & psAgencia
                            
    
   '************
   If psTipoC = 1 And psTipoM = 0 Then 'tipo = Comercial
       lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 1 & "' "
   End If
    
   If psTipoC = 0 And psTipoM = 1 Then 'tipo = MicroEmpresa
      lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 2 & "' "
   End If
   '******************
   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
    lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 1 & "' "
   End If
   If psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
    lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 2 & "' "
   End If

        
'On Error GoTo dError

   
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
       lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS X VENCER ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !Carta_Fianza & Space(1) & ImpreFormat(!Cliente, 30) & Space(2) _
                    & !Fecha_Emision & Space(2) & !Fecha_Vencimiento & ImpreFormat(!Importe, 12, 2, True) & Space(2) & ImpreFormat(!Acreedor, 40) _
                    & Space(2) & !Analista
                
                If Not OptAge Then
                    lsCadImp = lsCadImp & Space(2) & !NombAgencia
                End If
                lsCadImp = lsCadImp & Chr(10)

                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS REGISTRADOS ", "DEL " & Format(pdCMACT, "dd/mm/yyyy") & " - AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
       
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo148003_CartasFianzasVigentes_x_Vencer = lsCadBuffer

End Function
Public Function nRepo148004_CartasFianzaHonradas(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoC As Integer, ByRef psTipoM As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String) As String
        

Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

Dim P As String

lsSQL = "Select CF.cCtaCod as Carta_Fianza, CF.dAsignacion as Fecha_Emision, CF.dVencimiento as Fecha_Vencimiento," _
    & " P1.cPersNombre as Cliente, P2.cPersNombre as Acreedor, Pro.nSaldo as Importe," _
    & " RH.cUser as Analista, CF.dEjecutado  as Honrada, " _
    & " Ag.cAgeDescripcion as NombAgencia " _
    & " From ColocCartaFianza CF " _
    & " Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod " _
    & " Inner Join ProductoPersona PP  on CF.cCtaCod=PP.cCtaCod " _
    & " Inner Join ProductoPersona PP2 on CF.cCtaCod=PP2.cCtaCod " _
    & " Inner Join ProductoPersona PP3 on CF.cCtaCod=PP3.cCtaCod " _
    & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod " _
    & " Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod " _
    & " Inner Join RRHH    RH on RH.cPersCod = PP3.cPersCod " _
    & " Inner Join Agencias Ag on Ag.cAgeCod=Substring(CF.cCtaCod,4,2)" _
    & " WHERE PP.nPrdPersRelac = 20 and PP2.nPrdPersRelac=35 and PP3.nPrdPersRelac=28 " _
    & " and CF.dEjecutado between convert(datetime,'" & Format(pdDesde, "mm/dd/yyyy") & "') and  convert(datetime,'" & Format(pdHasta, "mm/dd/yyyy") & "')" _
    & " and CF.dEjecutado is not null "
    '
    If OptAge = True Then
        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)= '" & psAgencia & "'"
    Else
        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2) in (" & psListaAgencias & " )"
    End If
    'lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)=" & psAgencia
                            
    
   '************
   If psTipoC = 0 And psTipoM = 1 Then 'tipo = Comercial
       lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 2 & "' "
   End If
    
   If psTipoC = 1 And psTipoM = 0 Then 'tipo = MicroEmpresa
      lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 1 & "' "
   End If
   '******************
   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
    lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 1 & "' "
   End If
   If psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
    lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 2 & "' "
   End If


        
'On Error GoTo dError

   
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
       lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS  HONRADAS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !Carta_Fianza & Space(1) & ImpreFormat(!Cliente, 30) & Space(2) _
                    & !Fecha_Emision & Space(2) & !Fecha_Vencimiento & Space(2) & ImpreFormat(!Acreedor, 40) _
                    & Space(2) & !Analista & Space(7) & !Honrada
                
                If Not OptAge Then
                    lsCadImp = lsCadImp & Space(2) & !NombAgencia
                End If
                lsCadImp = lsCadImp & Chr(10)

                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS  HONRADAS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
       
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo148004_CartasFianzaHonradas = lsCadBuffer

End Function
Public Function nRepo148005_CartasFianzaDevueltas(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoC As Integer, ByRef psTipoM As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String) As String
        
        
        
Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

lsSQL = "Select Pro.cCtaCod as Carta_Fianza, CF.dAsignacion as Fecha_Emision, CF.dVencimiento as Fecha_Vencimiento," _
    & " P1.cPersNombre as Cliente, P2.cPersNombre as Acreedor, Pro.nSaldo as Importe," _
    & " RH.cUser as Analista, " _
    & " Ag.cAgeDescripcion as NombAgencia " _
    & " From Producto Pro   " _
    & " Inner Join ColocCartaFianza CF on CF.cCtaCod = Pro.cCtaCod " _
    & " Inner Join ProductoPersona PP  on CF.cCtaCod=PP.cCtaCod " _
    & " Inner Join ProductoPersona PP2 on CF.cCtaCod=PP2.cCtaCod " _
    & " Inner Join ProductoPersona PP3 on CF.cCtaCod=PP3.cCtaCod " _
    & " Inner Join ColocCFEstado   CFE on CF.cCtaCod = CFE.cCtaCod " _
    & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod " _
    & " Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod " _
    & " Inner Join RRHH    RH on RH.cPersCod = PP3.cPersCod " _
    & " Inner Join Agencias Ag on Ag.cAgeCod=Substring(CF.cCtaCod,4,2)" _
    & " WHERE PP.nPrdPersRelac = 20 and PP2.nPrdPersRelac=35 and PP3.nPrdPersRelac=28 " _
    & " and CFE.nPrdEstado = Pro.nPrdEstado and CFE.nPrdEstado = 2091 " _
    & " and CF.dEjecutado between convert (datetime , '" & Format(pdDesde, "mm/dd/yyyy") & "')" _
    & " and convert (datetime , '" & Format(pdHasta, "mm/dd/yyyy") & "') "
    
    If OptAge = True Then
        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)='" & psAgencia & "'"
    Else
        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2) in (" & psListaAgencias & " )"
    End If
    'lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)=" & psAgencia
   
   
    
   '************
   If psTipoC = 0 And psTipoM = 1 Then 'tipo = Comercial
        lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 2 & "' "
   End If
    
   If psTipoC = 1 And psTipoM = 0 Then 'tipo = MicroEmpresa
       lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 1 & "' "
   End If
    '******************
   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
      lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 1 & "' "
   End If
   If psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
     lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 2 & "' "
   End If
    


        
'On Error GoTo dError

   
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS DEVUELTAS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !Carta_Fianza & Space(1) & ImpreFormat(!Cliente, 30) & Space(2) _
                    & !Fecha_Emision & Space(2) & !Fecha_Emision & Space(2) & ImpreFormat(!Acreedor, 40) _
                    & Space(2) & !Analista
                
                If Not OptAge Then
                    lsCadImp = lsCadImp & Space(2) & !NombAgencia
                End If
                lsCadImp = lsCadImp & Chr(10)

                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS DEVUELTAS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
       
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo148005_CartasFianzaDevueltas = lsCadBuffer

End Function


Public Function nRepo148001_CartasFianzaEmitidas(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoC As Integer, ByRef psTipoM As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String) As String
        


Dim lsSQL As String
Dim lrDataRep As ADODB.Recordset
Dim loDataRep As dColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
'gColRelPersTitular
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String

lsSQL = "Select Pro.cCtaCod as Carta_Fianza, CF.dAsignacion as Fecha_Emision, CF.dVencimiento as Fecha_Vencimiento," _
    & " P1.cPersNombre as Cliente, P2.cPersNombre as Acreedor, Pro.nSaldo as Importe," _
    & " RH.cUser as Analista,  " _
    & "   ( Select C.cConsDescripcion as Describe from Constante C where CFE.nPrdEstado=C.nConsValor and nConsCod=3001) as Estado, "
    lsSQL = lsSQL & " Ag.cAgeDescripcion as NombAgencia "
    lsSQL = lsSQL & " From ColocCartaFianza CF  "
    lsSQL = lsSQL & " Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod "
    lsSQL = lsSQL & " Inner Join ProductoPersona PP  on CF.cCtaCod=PP.cCtaCod "
    lsSQL = lsSQL & " Inner Join ProductoPersona PP2 on CF.cCtaCod=PP2.cCtaCod "
    lsSQL = lsSQL & " Inner Join ProductoPersona PP3 on CF.cCtaCod=PP3.cCtaCod "
    lsSQL = lsSQL & " Inner Join ColocCFEstado   CFE on CF.cCtaCod = CFE.cCtaCod "
    lsSQL = lsSQL & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod "
    lsSQL = lsSQL & " Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod "
    lsSQL = lsSQL & " Inner Join RRHH    RH on RH.cPersCod = PP3.cPersCod "
    lsSQL = lsSQL & "Inner Join Agencias Ag on Ag.cAgeCod=Substring(CF.cCtaCod,4,2)"
    lsSQL = lsSQL & " WHERE PP.nPrdPersRelac = 20 and PP2.nPrdPersRelac=35 and PP3.nPrdPersRelac=28 "
    lsSQL = lsSQL & " and CFE.nPrdEstado = Pro.nPrdEstado "
    lsSQL = lsSQL & " and CF.dAsignacion between convert (datetime , '" & Format(pdDesde, "mm/dd/yyyy") & "')"
    lsSQL = lsSQL & " and convert (datetime , ' " & Format(pdHasta, "mm/dd/yyyy") & "') "
        
    If OptAge Then
        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2) in (" & psAgencia & ")"
    Else
        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2) in " & psListaAgencias
    End If
   
    
   '************
   If psTipoC = 0 And psTipoM = 1 Then 'tipo = Comercial
        lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 2 & "' "
   End If
    
   If psTipoC = 1 And psTipoM = 0 Then 'tipo = MicroEmpresa
       lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 1 & "' "
   End If
    '******************
   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
      lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 1 & "' "
   End If
   If psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
     lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 2 & "' "
   End If
    


        
'On Error GoTo dError

   
    Set loDataRep = New dColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS EMITIDAS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !Carta_Fianza & Space(1) & ImpreFormat(!Cliente, 30) & Space(2) _
                    & !Fecha_Emision & Space(2) & !Fecha_Vencimiento & Space(2) & ImpreFormat(!Acreedor, 40) _
                    & Space(2) & !Analista & Space(3) & ImpreFormat(!Estado, 12, 0)
                    
                If Not OptAge Then
                    lsCadImp = lsCadImp & Space(2) & !NombAgencia
                End If
                lsCadImp = lsCadImp & Chr(10)
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & Chr(12)
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS EMITIDAS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
       
        lsCadBuffer = lsCadBuffer & lsCadImp & Chr(12)
        
    End If
    nRepo148001_CartasFianzaEmitidas = lsCadBuffer

End Function



