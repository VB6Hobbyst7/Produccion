VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCartaFianzaReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim csNomCMAC As String
Dim csNomAgencia As String
Dim csCodUser As String
Dim csFechaSis As String
Public Function nRepoTarifario(Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim rs As New ADODB.Recordset
Dim CF As COMDCartaFianza.DCOMCartaFianza
Dim lsCadena As String
Dim Tabula As String
Dim X As String
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
Set CF = New COMDCartaFianza.DCOMCartaFianza
Set rs = CF.RecuperaCF_Tarifario
Tabula = Space(5)
lsCadena = String(80, "-") & oFunI.gPrnSaltoLinea
lsCadena = lsCadena & Tabula & "Cod. Tarif." & Space(2) & "Modalidad" & Space(20) & "Moneda" & Space(5) & "Tasa Trim." & Space(2) & "Monto Minimo" & oFunI.gPrnSaltoLinea
lsCadena = lsCadena & String(80, "-") & oFunI.gPrnSaltoLinea
While Not rs.EOF
    lsCadena = lsCadena & Tabula & ofun.ImpreFormat(rs!ctarifCod, 5) & Space(3) & ofun.ImpreFormat(rs!Modalidad, 25) & Space(5) & ofun.ImpreFormat(rs!Moneda, 7) & Space(3) & ofun.ImpreFormat(rs!nTasaTrim, 4, 2, False) & Space(4) & ofun.ImpreFormat(rs!nMontoMinimo, 8, 2, False) & oFunI.gPrnSaltoLinea
    rs.MoveNext
Wend
nRepoTarifario = lsCadena
Set rs = Nothing
Set CF = Nothing
End Function

'TODOCOMPLETA FALTA DEFINIR EL DRCARTAFIANZA**************************
'**********************************************************************
Public Function nImprimeCartaFianza(ByVal pnCodCta As String) As String
'Dim Rs As ADODB.Recordset
'Dim oCon As COMConecta.DCOMConecta 'COMConecta.DCOMConecta
'Dim sql  As String
'Set oCon = New COMConecta.DCOMConecta
'Set Rs = New ADODB.Recordset
'
'oCon.AbreConexion
'sql = "SELECT GETDATE() "
'Set Rs = oCon.CargaRecordSet(sql)
'oCon.CierraConexion
'Set oCon = Nothing
'With DRCartaFianza
'      Set .DataSource = Rs
'      .DataMember = ""
'      .WindowState = vbMaximized
'      .Title = "CARTA FIANZA"
'      .inicio pnCodCta
'      .Refresh
'      .Show vbModal
' End With
''DRCartaFianza.inicio pnCodCta
''DRCartaFianza.Show vbModal

End Function
Public Function nRepoDuplicado(ByVal pnCodCta As String, ByRef SoSu As Integer, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
Dim loRs As NCOMCartaFianzaValida
Dim lrDataT As New ADODB.Recordset
Dim lrDataCR As New ADODB.Recordset
Dim lrDataCF As New ADODB.Recordset

Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lnNegritaON As String
Dim lnNegritaOFF As String

Dim sCliente As String
Dim sAcreedor As String
Dim sAnalista As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr(27) & Chr$(70)
    
    Set loRs = New NCOMCartaFianzaValida
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(pnCodCta)
    Set lrDataT = loRs.RecuperaDatosT(pnCodCta)
    lsCadImp = ""
    If lrDataT Is Nothing Or (lrDataT.BOF And lrDataT.EOF) Then
        psmensaje = " No Existen Datos para el reporte"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1
        'CABECERA
        Select Case SoSu
            Case 2
            lsCadImp = lsCadImp & nRepoCabecera("R E G I S T R O   D E   S O L I C I T U D   D E    C A R T A   F I A N Z A", " ", lnPage, 130, "", 100)
            Case 3
            lsCadImp = lsCadImp & nRepoCabecera("R E S U M E N  D E  C O M I T E", " ", lnPage, 130, "", 100)
        End Select
        
        'lsCadImp = lsCadImp & nRepoCabecera("R E S U M E N   D E   C A R T A    F I A N Z A ", " ", lnPage, 130, "", 100)
        lnIndice = 0:  lnLineas = 7
        '***********************************************
        
        lsCadImp = lsCadImp & Space(20) & "CARTA FIANZA:" & Space(5) & lrDataCF!Carta
        lsCadImp = lsCadImp & Space(30) & lrDataCF!tipo & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(10) & ofun.ImpreFormat("  DATOS DEL TITULAR", 40, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 4
        '***************** DATOS TITULAR
        sCliente = lrDataT!Nombre
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("CLIENTE", 20, 0) & ":" & Space(2) & lrDataT!Nombre & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("DOC. IDENTIDAD", 20, 0) & ":" & Space(2) & lrDataT!DocDNI & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("DOC. TRIBUTARIO", 20, 0) & ":" & Space(2) & lrDataT!DocRUC & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("DIRECCION", 20, 0) & ":" & Space(2) & lrDataT!Direccion & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 4
        
        If SoSu = 3 Then
         Set lrDataCR = loRs.RecuperaDatosJuridica(lrDataT!Codigo)
         'MsgBox lrDataT!Codigo

         If lrDataCR.RecordCount = 0 Or (lrDataCR.EOF And lrDataCR.BOF) Then
         Else
         lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("TIEMPO CREACION", 20, 0) & ":" & Space(2) & lrDataT!Nacimiento & oFunI.gPrnSaltoLinea
         lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("TIPO JURIDICO", 20, 0) & ":" & Space(2) & "" & lrDataCR!DescribeJur & oFunI.gPrnSaltoLinea
         lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("MAGNITUD EMPRESA", 20, 0) & ":" & Space(2) & "" & lrDataCR!Magnitud & oFunI.gPrnSaltoLinea
         lnLineas = lnLineas + 3
         End If
        End If
        
        '--------------------------------------------------
        ' DATOS ACREEDOR
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(10) & ofun.ImpreFormat("  DATOS DEL ACREEDOR", 40, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 3
        
        Set lrDataCR = loRs.RecuperaDatosAcreedor(pnCodCta)
        If lrDataCR.EOF And lrDataCR.BOF Then
        Else
        sAcreedor = lrDataCR!Nombre
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("CLIENTE", 20, 0) & ":" & Space(2) & lrDataCR!Nombre & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("DOC. IDENTIDAD", 20, 0) & ":" & Space(2) & lrDataCR!DocDNI & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("DOC. TRIBUTARIO", 20, 0) & ":" & Space(2) & lrDataCR!DocRUC & oFunI.gPrnSaltoLinea 'WIOR 20130417 SE REEMPLAZO lrDataT POR lrDataCR
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("DIRECCION", 20, 0) & ":" & Space(2) & lrDataCR!Direccion & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 4
        End If
        '-------------------------------------------------
        
        'DATOS FUENTE DE INGRESO
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(10) & ofun.ImpreFormat("  FUENTE DE INGRESO", 40, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 3
        
        Set lrDataCR = loRs.RecuperaFuenteIngreso(lrDataT!Codigo)
        If lrDataCR.EOF And lrDataCR.BOF Then
        Else
         lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Tipo Ingreso", 20, 0) & ":" & Space(2) & lrDataCR!Tipo_Ingreso & oFunI.gPrnSaltoLinea
         lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Nombre/Razon Social", 20, 0) & ":" & Space(2) & lrDataCR!Razon & oFunI.gPrnSaltoLinea
         lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Direccion", 20, 0) & ":" & Space(2) & lrDataCR!Direccion & oFunI.gPrnSaltoLinea
         lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Actividad", 20, 0) & ":" & Space(2) & lrDataCR!Actividad & oFunI.gPrnSaltoLinea
         lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Ultima Actualizacion", 20, 0) & ":" & Format(lrDataCR!UltimaAct, "dd/mm/yyyy hh:mm:ss") & Space(2) & oFunI.gPrnSaltoLinea
         
         lnLineas = lnLineas + 5
        End If
        
        '--------------------------------------------------
        'DATOS CLIENTES RELACIONADOS
        'lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea  & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(10) & ofun.ImpreFormat("  CLIENTES RELACIONADOS", 40, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 2
        
        lsCadImp = lsCadImp & Space(4) & "CODIGO" & Space(10) & "NOMBRE" & Space(42) & "RELACIONES" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 2
        Set lrDataCR = loRs.RecuperaPersonaRelacion(pnCodCta)
        
        While Not lrDataCR.EOF
            lsCadImp = lsCadImp & Space(4) & lrDataCR!Codpersona
            lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat(lrDataCR!Nombre, 40, 0) & Space(10)
            lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat(lrDataCR!Relacion, 30, 0) & oFunI.gPrnSaltoLinea
            lrDataCR.MoveNext
            lnLineas = lnLineas + 1
         Wend
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 1
        '----------------------------------------------------
        'DATOS CARTA FIANZA
        
        'lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea  & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(10) & ofun.ImpreFormat("  DATOS DE LA CARTA FIANZA", 40, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 2
        
        sAnalista = lrDataCF!Analista
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Analista", 20, 0) & ":" & Space(2) & ofun.ImpreFormat(lrDataCF!Analista, 30, 0) & Space(10)
        lsCadImp = lsCadImp & ofun.ImpreFormat("Fuente de Ingreso", 20, 0) & ":" & Space(2) & lrDataCF!F_Ingreso & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Fecha Asignacion", 20, 0) & ":" & Space(2) & ofun.ImpreFormat(Format(lrDataCF!F_Asignacion, "dd/mm/yyyy"), 30, 0) & Space(10)
        'lsCadImp = lsCadImp & ofun.ImpreFormat("Modalidad", 20, 0) & ":" & Space(2) & lrDataCF!Modalidad & oFunI.gPrnSaltoLinea 'Comento JOEP20181221 CP
        lsCadImp = lsCadImp & ofun.ImpreFormat("Modalidad", 20, 0) & ":" & Space(2) & IIf(lrDataCF!nModalidad = 13, lrDataCF!OtrsModalidades, lrDataCF!Modalidad) & oFunI.gPrnSaltoLinea 'JOEP20181221 CP
        
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Fecha de Vigencia", 20, 0) & ":" & Space(2) & ofun.ImpreFormat(Format(lrDataCF!Vence, "dd/mm/yyyy"), 30, 0) & Space(10)
        lsCadImp = lsCadImp & ofun.ImpreFormat("Condicion", 20, 0) & ":" & Space(2) & lrDataCF!Condicion & oFunI.gPrnSaltoLinea
        
        If SoSu = 2 Then
            lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Monto Solicitado", 20, 0) & ":" & Space(2) & IIf(Mid(lrDataCF!Carta, 9, 1) = "1", "S/. ", "$ ") & Format(lrDataCF!nMonto, "##,##0.00") & oFunI.gPrnSaltoLinea
        End If
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 5
        
            '----------------------------------------------------------------
        'RELAC INTERNAAS
        
        If SoSu = 3 Then
        'lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(10) & ofun.ImpreFormat("  RELACIONES INTERNAS", 40, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 2
        Set lrDataCR = loRs.RecuperaRelacinesInternas(lrDataT!Codigo)
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("CLIENTE ", 30, 0) & Space(2)
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("RELACION ", 15, 0) & Space(4)
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("CARTA FIANZA ", 20, 0) & Space(2)
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("MONTO ", 10, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(130, "-") & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 2
        While Not lrDataCR.EOF
            lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat(lrDataCR!Cliente, 30, 0) & Space(2)
            lsCadImp = lsCadImp & ofun.ImpreFormat(lrDataCR!Relacion, 15, 0) & Space(2)
            lsCadImp = lsCadImp & ofun.ImpreFormat(lrDataCR!Carta, 20, 0) & Space(2)
            lsCadImp = lsCadImp & ofun.ImpreFormat(lrDataCR!Saldo, 10, 2) & Space(1) & lrDataCR!Moneda & oFunI.gPrnSaltoLinea
            lrDataCR.MoveNext
            lnLineas = lnLineas + 1
        Wend
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "OBSERVACIONES " & String(117, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(131, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(131, "-") & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(131, "-") & oFunI.gPrnSaltoLinea
        End If
        
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        'lsCadImp = lsCadImp & Space(5) & String(50, "-") & Space(10) & String(50, "-") & oFunI.gPrnSaltoLinea
        If SoSu = 2 Then
            lsCadImp = lsCadImp & Space(5) & String(50, "-") & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(23) & "CLIENTE" & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & Space(15) & sCliente & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        End If
        
        'lsCadImp = lsCadImp & Space(35) & String(50, "-") & oFunI.gPrnSaltoLinea
        'lsCadImp = lsCadImp & Space(55) & "ANALISTA" & oFunI.gPrnSaltoLinea
        'lsCadImp = lsCadImp & Space(45) & sAnalista & oFunI.gPrnSaltoLinea
    

        If lnLineas >= 55 Then
           lnPage = lnPage + 1
           lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
           Select Case SoSu
            Case 2
                lsCadImp = lsCadImp & nRepoCabecera("R E G I S T R O   D E   S O L I C I T U D   D E    C A R T A   F I A N Z A", "  ", lnPage, 130, "", 100)
            Case 3
                lsCadImp = lsCadImp & nRepoCabecera("R E S U M E N  D E  C O M I T E", " ", lnPage, 130, "", 100)
            End Select
           lnLineas = 7
        End If
       
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepoDuplicado = lsCadBuffer


End Function

Public Function nRepoCartaFianza_Estados(ByVal pnCodCta As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lrDataRep As New ADODB.Recordset
Dim lrDataT As New ADODB.Recordset
Dim lrDataH As New ADODB.Recordset
Dim lrDataG As New ADODB.Recordset
Dim lrDataO As New ADODB.Recordset

'Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim loRs As NCOMCartaFianzaValida

Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lnNegritaON As String
Dim lnNegritaOFF As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    
    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr(27) & Chr$(70)
    
    
    Set loRs = New NCOMCartaFianzaValida
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
        lsCadImp = lsCadImp & "Nª Carta Fianza : " & lrDataRep!Carta & Space(15) & "Tipo : " & lrDataRep!tipo & _
                    Space(15) & "Estado Actual : " & lrDataRep!Estado & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & lnNegritaOFF
        '***********************************************
        
        lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Space(10) & ofun.ImpreFormat("  DATOS GENERALES", 40, 0) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Analista", 20, 0) & ":" & Space(2) & ofun.ImpreFormat(lrDataRep!Analista, 30, 0) & Space(10)
        lsCadImp = lsCadImp & ofun.ImpreFormat("Fuente de Ingreso", 20, 0) & ":" & Space(2) & lrDataRep!F_Ingreso & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Apoderado", 20, 0) & ":" & Space(2) & ofun.ImpreFormat(IIf(IsNull(lrDataRep!Apoderado), "", lrDataRep!Apoderado), 30, 0) & Space(10)
        lsCadImp = lsCadImp & ofun.ImpreFormat("Fecha de Vigencia", 20, 0) & ":" & Space(2) & Format(lrDataRep!Vence, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
        
        'lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Modalidad", 20, 0) & ":" & Space(2) & ofun.ImpreFormat(lrDataRep!Modalidad, 30, 0) & Space(10)'comento JOEP20181227 CP
        lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat("Modalidad", 20, 0) & ":" & Space(2) & ofun.ImpreFormat(IIf(lrDataRep!nModalidad = 13, lrDataRep!OtrsModalidades, lrDataRep!Modalidad), 30, 0) & Space(10)
        
        lsCadImp = lsCadImp & ofun.ImpreFormat("Condicion", 20, 0) & ":" & Space(2) & lrDataRep!Condicion & oFunI.gPrnSaltoLinea
        
        lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 7
        
        Set lrDataT = loRs.RecuperaPersonaRelacion(pnCodCta)
         lsCadImp = lsCadImp & Space(10) & "NOMBRE" & Space(37) & "RELACIONES" & oFunI.gPrnSaltoLinea
         lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
         lnLineas = lnLineas + 2
         
         While Not lrDataT.EOF
            lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat(lrDataT!Nombre, 40, 0) & Space(10)
            lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat(lrDataT!Relacion, 30, 0) & oFunI.gPrnSaltoLinea
            lrDataT.MoveNext
            lnLineas = lnLineas + 1
         Wend
        lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        
        '-------------------------------
        Set lrDataRep = loRs.RecuperaHistorial(pnCodCta)
        Set lrDataH = loRs.RecuperaHistorial(pnCodCta)
        
        lsCadImp = lsCadImp & Space(53) & "FECHA" & Space(28) & "MONTO" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
         lnLineas = lnLineas + 3
         While Not lrDataH.EOF
            lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat(lrDataH!Estado, 10, 0)
            lsCadImp = lsCadImp & Space(42) & Format(lrDataH!Fecha, "dd/mm/yyyy")
            lsCadImp = lsCadImp & Space(10) & ofun.ImpreFormat(lrDataH!Monto, 15, 2) & oFunI.gPrnSaltoLinea
            lrDataH.MoveNext
            lnLineas = lnLineas + 1
         Wend
         
         lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
         lnLineas = lnLineas + 2
        
         '-------------------------------
         
         Set lrDataG = loRs.RecuperaGarantias(pnCodCta)
         
        lsCadImp = lsCadImp & Space(2) & "Nª Garantia" & Space(12)
        lsCadImp = lsCadImp & "Tipo Garantia" & Space(15) & "Descripcion" & Space(21) & "Documento" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
        
        
        lnLineas = lnLineas + 3
        While Not lrDataG.EOF
            lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat(lrDataG!NroGarantia, 10, 0)
            lsCadImp = lsCadImp & Space(13) & Format(lrDataG!Tipo_Garn, "dd/mm/yyyy")
            lsCadImp = lsCadImp & Space(20) & ofun.ImpreFormat(lrDataG!Descripcion, 20, 0)
            lsCadImp = lsCadImp & Space(12) & ofun.ImpreFormat(lrDataG!Tipo_Doc, 15, 0) & oFunI.gPrnSaltoLinea
            lrDataG.MoveNext
            lnLineas = lnLineas + 1
        Wend
        lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
        lnLineas = lnLineas + 2
        
        
        '-----------------------------------------------
        Set lrDataO = loRs.RepcuperaOtrosDt(pnCodCta)
        If lrDataO Is Nothing Or (lrDataO.EOF And lrDataO.BOF) Then
        Else
         lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
         lsCadImp = lsCadImp & lnNegritaON
         lsCadImp = lsCadImp & "Motivo Rechazo :" & ofun.ImpreFormat(lrDataO!Motivo, 20, 0) & Space(2)
         lsCadImp = lsCadImp & "Fecha :" & Format(lrDataO!Fecha, "dd/mm/yyyy") & oFunI.gPrnSaltoLinea
         lsCadImp = lsCadImp & lnNegritaOFF
         lsCadImp = lsCadImp & String(150, "-") & oFunI.gPrnSaltoLinea
         lnLineas = lnLineas + 3
        End If
        
        '----------------------------------------------------
        If lnLineas >= 55 Then
           lnPage = lnPage + 1
           lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
           lsCadImp = lsCadImp & nRepoCabecera("R E S U M E N   D E   C A R T A    F I A N Z A ", " ", lnPage, 150, "", 100)
           lnLineas = 7
        End If
       
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    '-End if
    nRepoCartaFianza_Estados = lsCadBuffer

Set oFunI = Nothing

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
        ByVal pnCodReporte As Long, Optional ByVal psImpresora As Impresoras = gEPSON) As String
        
Dim lsCadImp As String
Dim loImprimeCab As COMNColoCPig.NCOMColPImpre
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    
    Set loImprimeCab = New COMNColoCPig.NCOMColPImpre
        lsCadImp = loImprimeCab.nImprimeCabeceraReportes(csNomCMAC, csNomAgencia, csCodUser, csFechaSis, psTitulo, psSubTitulo, pnPagina, pnAnchoLinea, psComenta)
    Set loImprimeCab = Nothing
    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
    Select Case pnCodReporte
        Case 148001  ' Cartas Fianzas Aprobadas
            lsCadImp = lsCadImp & Space(1) & "Item   CartaFianza         Cliente                        Emision  Vencimiento    Acreedor                                Analista   Estado     Tip_Credito" & oFunI.gPrnSaltoLinea
        Case 148002 ' Cartas Fianzas vigentes, DAOR 20070111
            lsCadImp = lsCadImp & Space(1) & "CartaFianza       #Ren     Cliente                       Tipo de crédito        Emision Vencimiento  #C/Cont      Importe      Beneficiario                   Modalidad   Agencia" & oFunI.gPrnSaltoLinea
'        Case 148002 ' Cartas Fianzas Emitidas
'            lsCadImp = lsCadImp & Space(1) & "Item   CartaFianza           Cliente                        Emision  Vencimiento    Acreedor                                Analista " & oFunI.gPrnSaltoLinea
        Case 148003 ' Cartas Fianzas Vigentes
            lsCadImp = lsCadImp & Space(1) & "Item   CartaFianza           Cliente                        Emision  Vencimiento      Importe    Acreedor                                 Analista  Tip_Credito" & oFunI.gPrnSaltoLinea
        Case 148004 ' Cartas Fianzas Vencidas
            lsCadImp = lsCadImp & Space(1) & "Item   CartaFianza           Cliente                        Emision  Vencimiento    Acreedor                                Analista   Carta Honrada Tip_Credito" & oFunI.gPrnSaltoLinea
        Case 148005 ' Cartas Fianzas Devueltas
            lsCadImp = lsCadImp & Space(1) & "Item   CartaFianza           Cliente                        Emision  Vencimiento    Acreedor                                Analista   Tip_Credito" & oFunI.gPrnSaltoLinea
        Case 148006 ' Cartas Fianzas vigentes, By Capi Acta 035-2007
            lsCadImp = lsCadImp & Space(1) & "CartaFianza       #Ren     Cliente                    Emision Vencimiento Cancelacion  #C/Cont      Importe      Beneficiario                  Modalidad  Agencia   Tip_Credito" & oFunI.gPrnSaltoLinea
    End Select
    lsCadImp = lsCadImp & String(pnAnchoLinea, "-") & oFunI.gPrnSaltoLinea
nRepoCabecera = lsCadImp
Set oFunI = Nothing
End Function

'**DAOR 20070111, se adecuó al formato de la CAJA Maynas
Public Function nRepo148002_CartasFianzaVigentes(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdCMACT As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psmensaje As String, Optional psImpresora As Impresoras = gEPSON _
  , Optional ByRef dFecIni As String, Optional ByRef dFecFin As String) As String
        '***RECO 2013-07-18:SE AGREGO DOS PARAMETROS NUEVOS (dFecIni,dFecFin)*
Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim ofun As New COMFunciones.FCOMImpresion

Dim P As String, sAges As String, sTipo As String, sMone As String
Dim lcAge As String, lcMone As String, lnTot As Double

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora


'    sCadAge = ""
'    For I = 0 To UBound(pMatAgencias) - 1
'        sCadAge = sCadAge & pMatAgencias(I) & ","
'    Next I
'    sCadAge = Mid(sCadAge, 1, Len(sCadAge) - 1)



''*** PEAC 20090401

'lsSQL = "Select CF.cCtaCod as Carta_Fianza, CF.dAsignacion as Fecha_Emision, CF.dVencimiento as Fecha_Vencimiento," _
'    & " P1.cPersNombre as Cliente, P2.cPersNombre as Acreedor, CC.nMontoCol as Importe,isnull(CF.nRenovacion,0) as Num_Renovacion, " _
'    & " RH.cUser as Analista, CFP.nPoliza as Num_Poliza, " _
'    & " Ag.cAgeDescripcion as NombAgencia, " _
'    & " ( Select C.cConsDescripcion  from Constante C where Pro.nPrdEstado=C.nConsValor and nConsCod=3001) as Estado, " _
'    & " ( select C.cConsDescripcion from Constante C where C.nConsValor=CF.nModalidad and nConsCod=" & gColCFModalidad & ") as Modalidad " _
'    & " From Producto Pro " _
'    & " Inner Join colocaciones CC on CC.cCtaCod = Pro.cCtaCod " _
'    & " " _
'    & " Inner Join ColocCartaFianza CF on CF.cCtaCod = Pro.cCtaCod " _
'    & " Inner Join ColocCFPoliza CFP on CF.cCtaCod = CFP.cCtaCod " _
'    & " Inner Join ProductoPersona PP  on CF.cCtaCod=PP.cCtaCod " _
'    & " Inner Join ProductoPersona PP2 on CF.cCtaCod=PP2.cCtaCod " _
'    & " Inner Join ProductoPersona PP3 on CF.cCtaCod=PP3.cCtaCod " _
'    & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod " _
'    & " Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod " _
'    & " Inner Join RRHH    RH on RH.cPersCod = PP3.cPersCod " _
'    & " Inner Join Agencias Ag on Ag.cAgeCod=Substring(CF.cCtaCod,4,2)" _
'    & " WHERE PP.nPrdPersRelac = 20 and PP2.nPrdPersRelac=35 and PP3.nPrdPersRelac=28 " _
'    & " And Pro.nPrdEstado in (" & gColocEstVigNorm & ", " & gColocEstRenovada & ")" _
'    & "  "
'*** FIN PEAC
    
'lsSQL = "Select CF.cCtaCod as Carta_Fianza, CF.dAsignacion as Fecha_Emision, CF.dVencimiento as Fecha_Vencimiento," _
'    & " P1.cPersNombre as Cliente, P2.cPersNombre as Acreedor, CC.nMontoCol as Importe, " _
'    & " RH.cUser as Analista, " _
'    & " Ag.cAgeDescripcion as NombAgencia, " _
'    & " ( Select C.cConsDescripcion  from Constante C where Pro.nPrdEstado=C.nConsValor and nConsCod=3001) as Estado, " _
'    & " ( select C.cConsDescripcion from Constante C where C.nConsValor=CF.nModalidad and nConsCod=" & gColCFModalidad & ") as Modalidad " _
'    & " From Producto Pro " _
'    & " Inner Join colocaciones CC on CC.cCtaCod = Pro.cCtaCod " _
'    & " " _
'    & " Inner Join ColocCartaFianza CF on CF.cCtaCod = Pro.cCtaCod " _
'    & " Inner Join ProductoPersona PP  on CF.cCtaCod=PP.cCtaCod " _
'    & " Inner Join ProductoPersona PP2 on CF.cCtaCod=PP2.cCtaCod " _
'    & " Inner Join ProductoPersona PP3 on CF.cCtaCod=PP3.cCtaCod " _
'    & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod " _
'    & " Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod " _
'    & " Inner Join RRHH    RH on RH.cPersCod = PP3.cPersCod " _
'    & " Inner Join Agencias Ag on Ag.cAgeCod=Substring(CF.cCtaCod,4,2)" _
'    & " WHERE PP.nPrdPersRelac = 20 and PP2.nPrdPersRelac=35 and PP3.nPrdPersRelac=28 " _
'    & " And Pro.nPrdEstado=" & gColocEstVigNorm & "" _
'    & "  "

''*** PEAC 20090401
'    If OptAge = True Then
'        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)='" & psAgencia & "'"
'    Else
'        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2) in (" & psListaAgencias & " )"
'    End If
    'lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)=" & psAgencia
'*** FIN PEAC

    If OptAge = True Then
        sAges = psAgencia
    Else
        sAges = Replace(Replace(psListaAgencias, "'", ""), " ", "")
    End If
    
  
'   If psTipoC = 0 And psTipoM = 1 Then
'       sTipo = "2"
'   ElseIf psTipoC = 1 And psTipoM = 0 Then
'      sTipo = "1"
'    Else
'      sTipo = "1,2"
'   End If

'*** BRGO BASILEA II
    If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
        sTipo = "1,2,3,4,5"
    Else
        If psTipoCor = 1 Then
          sTipo = sTipo & ",1"
        End If
        If psTipoGraEmp = 1 Then
          sTipo = sTipo & ",2"
        End If
        If psTipoMedEmp = 1 Then
          sTipo = sTipo & ",3"
        End If
        If psTipoPeqEmp = 1 Then
          sTipo = sTipo & ",4"
        End If
        If psTipoMicEmp = 1 Then
          sTipo = sTipo & ",5"
        End If
        sTipo = Mid(sTipo, 2, Len(sTipo) - 1)
    End If
'**********************
   


   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
    sMone = "1"
   ElseIf psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
    sMone = "2"
    Else
    sMone = "1,2"
   End If
        
   '**************************
        
'On Error GoTo dError

    '*** PEAC 20090401
    '****RECO 2013/07/18*****
    lsSql = " exec stp_sel_ReporteCartasFianzasVigentes '" & sAges & "','" & sMone & "','" & sTipo & "','" & dFecIni & "','" & dFecFin & "'"
    '****END RECO************

    'lsSQL = " exec stp_sel_ReporteCartasFianzasVigentes '" & sAges & "','" & sMone & "','" & sTipo & "'"

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia "
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 179, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
        
'            Do While Not lrDataRep.EOF
'                lnIndice = lnIndice + 1
'                lnLineas = lnLineas + 1
'
'                'Cadena de Impresion
'                'lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & ofun.ImpreFormat(!Carta_Fianza, 18) & Space(1) & "999" & Space(1) & ofun.ImpreFormat(!Cliente, 30) & Space(1)
'                lsCadImp = lsCadImp & ofun.ImpreFormat(!Carta_Fianza, 18, 0) & Space(1) & ofun.ImpreFormat(!Num_Renovacion, 3, 0) & Space(1) & ofun.ImpreFormat(!Cliente, 30, 0) & Space(1) _
'                    & !Fecha_Emision & Space(1) & !Fecha_Vencimiento & Space(1) & Format(!Num_Poliza, "0000000") & Space(1) & ofun.ImpreFormat(!Importe, 12, 2, True) & Space(1) & ofun.ImpreFormat(!Acreedor, 30, 0) _
'                    & Space(1) & ofun.ImpreFormat(!Modalidad, 12, 0)
'
'                If Not OptAge Then
'                    lsCadImp = lsCadImp & Space(1) & ofun.ImpreFormat(!NombAgencia, 12, 0)
'                End If
'                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'
'
'                If lnIndice Mod 300 = 0 Then
'                    lsCadBuffer = lsCadBuffer & lsCadImp
'                    lsCadImp = ""
'                End If
'
'                If lnLineas >= 55 Then
'                    lnPage = lnPage + 1
'                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
'                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
'                    lnLineas = 7
'                End If
'                .MoveNext
'            Loop


'    Do While Not R.EOF
'        sCadImp = sCadImp & "AGENCIA : " & R!Agencia & lnSaltoLinDoc
'        vage = R!CodAge
'
'        J = 0
'
'        Do While R!CodAge = vage
'            I = I + 1
'
'            J = J + 1
'            sCadImp = sCadImp & oImpre.ImpreFormat(J, 4, 0)
'            sCadImp = sCadImp & oImpre.ImpreFormat(R!cCtaCod, 18)
            
            
            Do While Not lrDataRep.EOF
            
                lsCadImp = lsCadImp & "Agencia : " & !NombAgencia & oFunI.gPrnSaltoLinea
                lcAge = !cCodAge
                                
                Do While lrDataRep!cCodAge = lcAge
                    lsCadImp = lsCadImp & "Moneda : " & IIf(!cCodMone = "1", "SOLES", "DOLAR") & oFunI.gPrnSaltoLinea
                    lcMone = !cCodMone
                                
                    lnTot = 0
                    Do While lrDataRep!cCodMone = lcMone And lrDataRep!cCodAge = lcAge
            
                        lnIndice = lnIndice + 1
                        lnLineas = lnLineas + 1

                        lsCadImp = lsCadImp & ofun.ImpreFormat(!Carta_Fianza, 18, 0) & Space(1) & ofun.ImpreFormat(!Num_Renovacion, 3, 0) & Space(1) & ofun.ImpreFormat(!Cliente, 30, 0) & Space(1) & ofun.ImpreFormat(!TpoCred, 24, 0) & Space(1) _
                            & !Fecha_Emision & Space(1) & !Fecha_Vencimiento & Space(1) & Format(!Num_Poliza, "0000000") & Space(1) & ofun.ImpreFormat(!Importe, 12, 2, True) & Space(1) & ofun.ImpreFormat(!Acreedor, 30, 0) _
                            & Space(1) & ofun.ImpreFormat(!Modalidad, 12, 0)
                            
                            If Not OptAge Then
                                lsCadImp = lsCadImp & Space(1) & ofun.ImpreFormat(!NombAgencia, 12, 0)
                            End If
                            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                
                            If lnIndice Mod 300 = 0 Then
                                lsCadBuffer = lsCadBuffer & lsCadImp
                                lsCadImp = ""
                            End If
                
                            lnTot = lnTot + !Importe
                
                            lrDataRep.MoveNext
                            

                            
                            If lrDataRep.EOF Then
                                Exit Do
                            End If
                
                            If lnLineas >= 55 Then
                                lnPage = lnPage + 1
                                lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                                lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
                                lnLineas = 7
                            End If
                
                    Loop
                    
                    lsCadImp = lsCadImp & String(147, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & "Totales :" & Space(75) & ofun.ImpreFormat(lnTot, 12, 2, True)
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    If lrDataRep.EOF Then
                        Exit Do
                    End If
                Loop
                If lrDataRep.EOF Then
                    Exit Do
                End If
            Loop
            
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Total de Cartas Fianzas Vigentes: " & ofun.ImpreFormat(lnIndice, 5, 0)
       
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo148002_CartasFianzaVigentes = lsCadBuffer
Set oFunI = Nothing
End Function

'Public Function nRepo148002_CartasFianzaVigentes(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
'        ByVal pdCMACT As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
'        ByRef psMonedaE As Integer, ByRef psTipoC As Integer, ByRef psTipoM As Integer, _
'        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psmensaje As String, Optional psImpresora As Impresoras = gEPSON) As String
'
'Dim lsSQL As String
'Dim lrDataRep As New ADODB.Recordset
'Dim loDataRep As COMDColocPig.DCOMColPFunciones
'Dim lsCadImp As String
'Dim lsCadBuffer As String
'
'Dim lnIndice As Long
'Dim lnLineas As Integer
'Dim lnPage As Integer
'Dim lsOperaciones As String
'Dim ofun As New COMFunciones.FCOMImpresion
'
'Dim P As String
'
'Dim oFunI As New COMFunciones.FCOMVarImpresion
'oFunI.Inicia psImpresora
'
'
'lsSQL = "Select CF.cCtaCod as Carta_Fianza, CF.dAsignacion as Fecha_Emision, CF.dVencimiento as Fecha_Vencimiento," _
'    & " P1.cPersNombre as Cliente, P2.cPersNombre as Acreedor, CC.nMontoCol as Importe," _
'    & " RH.cUser as Analista, " _
'    & " Ag.cAgeDescripcion as NombAgencia, " _
'    & " ( Select C.cConsDescripcion  from Constante C where Pro.nPrdEstado=C.nConsValor and nConsCod=3001) as Estado " _
'    & " From Producto Pro " _
'    & " Inner Join colocaciones CC on CC.cCtaCod = Pro.cCtaCod " _
'    & " " _
'    & " Inner Join ColocCartaFianza CF on CF.cCtaCod = Pro.cCtaCod " _
'    & " Inner Join ProductoPersona PP  on CF.cCtaCod=PP.cCtaCod " _
'    & " Inner Join ProductoPersona PP2 on CF.cCtaCod=PP2.cCtaCod " _
'    & " Inner Join ProductoPersona PP3 on CF.cCtaCod=PP3.cCtaCod " _
'    & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod " _
'    & " Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod " _
'    & " Inner Join RRHH    RH on RH.cPersCod = PP3.cPersCod " _
'    & " Inner Join Agencias Ag on Ag.cAgeCod=Substring(CF.cCtaCod,4,2)" _
'    & " WHERE PP.nPrdPersRelac = 20 and PP2.nPrdPersRelac=35 and PP3.nPrdPersRelac=28 " _
'    & " And Pro.nPrdEstado=" & gColocEstVigNorm & "" _
'    & "  "
'
'    If OptAge = True Then
'        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)='" & psAgencia & "'"
'    Else
'        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2) in (" & psListaAgencias & " )"
'    End If
'    'lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)=" & psAgencia
'
'
'   '************
'   If psTipoC = 0 And psTipoM = 1 Then 'tipo = Comercial
'       lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 2 & "' "
'   End If
'
'   If psTipoC = 1 And psTipoM = 0 Then 'tipo = MicroEmpresa
'      lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 1 & "' "
'   End If
'   '******************
'   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
'    lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 1 & "' "
'   End If
'   If psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
'    lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 2 & "' "
'   End If
'
''On Error GoTo dError
'
'    Set loDataRep = New COMDColocPig.DCOMColPFunciones
'        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
'    Set loDataRep = Nothing
'
'    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
'        psmensaje = " No Existen Datos para el reporte en la Agencia "
'        Exit Function
'    Else
'        lnLineas = 0
'        lnPage = 1
'
'        'CABECERA
'        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
'
'        lnIndice = 0:  lnLineas = 7
'
'        With lrDataRep
'            Do While Not lrDataRep.EOF
'                lnIndice = lnIndice + 1
'                lnLineas = lnLineas + 1
'
'                'Cadena de Impresion
'                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !Carta_Fianza & Space(1) & ofun.ImpreFormat(!Cliente, 30) & Space(2) _
'                    & !Fecha_Emision & Space(2) & !Fecha_Vencimiento & Space(2) & ofun.ImpreFormat(!Acreedor, 40) _
'                    & Space(2) & !Analista
'
'                If Not OptAge Then
'                    lsCadImp = lsCadImp & Space(2) & !NombAgencia
'                End If
'                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'
'
'                If lnIndice Mod 300 = 0 Then
'                    lsCadBuffer = lsCadBuffer & lsCadImp
'                    lsCadImp = ""
'                End If
'
'                If lnLineas >= 55 Then
'                    lnPage = lnPage + 1
'                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
'                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
'                    lnLineas = 7
'                End If
'                .MoveNext
'            Loop
'        End With
'
'
'        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
'
'    End If
'    nRepo148002_CartasFianzaVigentes = lsCadBuffer
'Set oFunI = Nothing
'End Function

Public Function nRepo148003_CartasFianzasVigentes_x_Vencer(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, ByRef pdCMACT As Date, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
        

Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim sTipo As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

Dim P As String

'*** BRGO BASILEA II
    If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
        sTipo = "1,2,3,4,5"
    Else
        If psTipoCor = 1 Then
          sTipo = sTipo & ",1"
        End If
        If psTipoGraEmp = 1 Then
          sTipo = sTipo & ",2"
        End If
        If psTipoMedEmp = 1 Then
          sTipo = sTipo & ",3"
        End If
        If psTipoPeqEmp = 1 Then
          sTipo = sTipo & ",4"
        End If
        If psTipoMicEmp = 1 Then
          sTipo = sTipo & ",5"
        End If
        sTipo = Mid(sTipo, 2, Len(sTipo) - 1)
    End If
'**********************


lsSql = "Select CF.cCtaCod as Carta_Fianza, CF.dAsignacion as Fecha_Emision, CF.dVencimiento as Fecha_Vencimiento," _
    & " P1.cPersNombre as Cliente, P2.cPersNombre as Acreedor, Pro.nSaldo as Importe," _
    & " RH.cUser as Analista, Pro.dPrdEstado, " _
    & " Ag.cAgeDescripcion as NombAgencia " _
    & " ,K.cConsDescripcion Tip_Credito " _
    & " From ColocCartaFianza CF " _
    & " Inner Join Colocaciones C on C.cCtaCod = CF.cCtaCod " _
    & " Inner Join ColocCFEstado CFE on CFE.cCtaCod = CF.cCtaCod " _
    & " Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod " _
    & " Inner Join ProductoPersona PP  on CF.cCtaCod=PP.cCtaCod " _
    & " Inner Join ProductoPersona PP2 on CF.cCtaCod=PP2.cCtaCod " _
    & " Inner Join ProductoPersona PP3 on CF.cCtaCod=PP3.cCtaCod " _
    & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod " _
    & " Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod " _
    & " Inner Join RRHH    RH on RH.cPersCod = PP3.cPersCod " _
    & " Inner Join Agencias Ag on Ag.cAgeCod = C.cAgeCodAct" _
    & " Inner Join Constante K on K.nConsValor=C.cTpoCredCod and K.nConsCod = 3034 " _
    & " WHERE PP.nPrdPersRelac = 20 and PP2.nPrdPersRelac=35 and PP3.nPrdPersRelac=28 and CFE.nPrdEstado in (2020,2092) " _
    & "  " _
    & " and CF.dVencimiento between convert(datetime,'" & Format(pdDesde, "mm/dd/yyyy") & "') and  convert(datetime,'" & Format(pdHasta, "mm/dd/yyyy") & "')"
    
    If OptAge = True Then
        lsSql = lsSql & " and C.cAgeCodAct = '" & psAgencia & "'"
    Else
        lsSql = lsSql & " and C.cAgeCodAct in (" & psListaAgencias & " )"
    End If
                            
    
   '************
'   If psTipoC = 1 And psTipoM = 0 Then 'tipo = Comercial
'       lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 1 & "' "
'   End If
'
'   If psTipoC = 0 And psTipoM = 1 Then 'tipo = MicroEmpresa
'      lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 2 & "' "
'   End If

    If Len(sTipo) > 0 Then
      lsSql = lsSql & " and substring(C.cTpoCredCod,1,1) In (" & sTipo & ")"
    End If
   '******************
   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
    lsSql = lsSql & " and substring(CF.cCtaCod,9,1)='" & 1 & "' "
   End If
   If psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
    lsSql = lsSql & " and substring(CF.cCtaCod,9,1)='" & 2 & "' "
   End If

        
'On Error GoTo dError

   
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
       lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS X VENCER ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 160, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !Carta_Fianza & Space(1) & ofun.ImpreFormat(!Cliente, 30) & Space(2) _
                    & !Fecha_Emision & Space(2) & !Fecha_Vencimiento & ofun.ImpreFormat(!Importe, 12, 2, True) & Space(2) & ofun.ImpreFormat(!Acreedor, 40) _
                    & Space(2) & !Analista
                
                If Not OptAge Then
                    lsCadImp = lsCadImp & Space(2) & !NombAgencia
                End If
                lsCadImp = lsCadImp & Space(2) & ofun.ImpreFormat(!Tip_Credito, 11, 0) 'madm 20101115

                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CONTRATOS REGISTRADOS ", "DEL " & Format(pdCMACT, "dd/mm/yyyy") & " - AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 120, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
       
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo148003_CartasFianzasVigentes_x_Vencer = lsCadBuffer
    
Set oFunI = Nothing
End Function
Public Function nRepo148004_CartasFianzaHonradas(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
        

Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lcAge As String
Dim lcTipoM As String
Dim lcMone As String


Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

Dim P As String

    If OptAge = True Then
        lcAge = Replace(Replace(psAgencia, " ", ""), "'", "")
    Else
        lcAge = Replace(Replace(psListaAgencias, " ", ""), "'", "")
    End If
    
  
'*** BRGO BASILEA II
    If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
        lcTipoM = "1,2,3,4,5"
    Else
        If psTipoCor = 1 Then
          lcTipoM = lcTipoM & ",1"
        End If
        If psTipoGraEmp = 1 Then
          lcTipoM = lcTipoM & ",2"
        End If
        If psTipoMedEmp = 1 Then
          lcTipoM = lcTipoM & ",3"
        End If
        If psTipoPeqEmp = 1 Then
          lcTipoM = lcTipoM & ",4"
        End If
        If psTipoMicEmp = 1 Then
          lcTipoM = lcTipoM & ",5"
        End If
        lcTipoM = Mid(lcTipoM, 2, Len(lcTipoM) - 1)
    End If
'**********************
   
  
    If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
        lcMone = "1"
    ElseIf psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
        lcMone = "2"
    Else
        lcMone = "1,2"
    End If

   '******************

    lsSql = " exec stp_sel_CartasFianzaHonradas '" & Format(pdDesde, "yyyymmdd") & "','" & Format(pdHasta, "yyyymmdd") & "','" & lcAge & "','" & lcTipoM & "','" & lcMone & "'"
        

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
       lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS  HONRADAS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 162, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                'Cadena de Impresion
                'MADM 20101115
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !Carta_Fianza & Space(1) & ofun.ImpreFormat(!Cliente, 30) & Space(2) _
                    & !Fecha_Emision & Space(2) & !Fecha_Vencimiento & Space(2) & ofun.ImpreFormat(!Acreedor, 40) _
                    & Space(2) & !Analista & Space(7) & !Honrada & Space(1) & ofun.ImpreFormat(!Tip_Credito, 11, 0)
                
                If Not OptAge Then
                    lsCadImp = lsCadImp & Space(2) & !NombAgencia
                End If
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS  HONRADAS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 162, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
       
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo148004_CartasFianzaHonradas = lsCadBuffer
 Set oFunI = Nothing
End Function
Public Function nRepo148005_CartasFianzaDevueltas(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
        
Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim lcTipoM As String
Dim cMoneda As String
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

'*** BRGO BASILEA II
    If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
        lcTipoM = "1,2,3,4,5"
    Else
        If psTipoCor = 1 Then
          lcTipoM = lcTipoM & ",1"
        End If
        If psTipoGraEmp = 1 Then
          lcTipoM = lcTipoM & ",2"
        End If
        If psTipoMedEmp = 1 Then
          lcTipoM = lcTipoM & ",3"
        End If
        If psTipoPeqEmp = 1 Then
          lcTipoM = lcTipoM & ",4"
        End If
        If psTipoMicEmp = 1 Then
          lcTipoM = lcTipoM & ",5"
        End If
        
        lcTipoM = Mid(lcTipoM, 2, Len(lcTipoM) - 1)
    End If
    
    If OptAge = True Then
        psListaAgencias = psAgencia
    End If
    
    If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
        cMoneda = "1"
    ElseIf psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
        cMoneda = "2"
    Else
        cMoneda = "1,2"
   End If
    lsSql = "stp_sel_CF_Devueltas '" & Format(pdDesde, "yyyy/MM/dd") & "','" & Format(pdHasta, "yyyy/MM/dd") & "','" & psListaAgencias & "','" & cMoneda & "','" & lcTipoM & "'"
    
'    lsSQL = "Select Pro.cCtaCod as Carta_Fianza, CF.dAsignacion as Fecha_Emision, CF.dVencimiento as Fecha_Vencimiento," _
'    & " P1.cPersNombre as Cliente, P2.cPersNombre as Acreedor, Pro.nSaldo as Importe," _
'    & " RH.cUser as Analista, " _
'    & " Ag.cAgeDescripcion as NombAgencia " _
'    & " ,K.cConsDescripcion Tip_Credito " _
'    & " From Producto Pro   " _
'    & " Inner Join Colocaciones C on C.cCtaCod = Pro.cCtaCod " _
'    & " Inner Join ColocCartaFianza CF on CF.cCtaCod = Pro.cCtaCod " _
'    & " Inner Join ProductoPersona PP  on CF.cCtaCod=PP.cCtaCod " _
'    & " Inner Join ProductoPersona PP2 on CF.cCtaCod=PP2.cCtaCod " _
'    & " Inner Join ProductoPersona PP3 on CF.cCtaCod=PP3.cCtaCod " _
'    & " Inner Join ColocCFEstado   CFE on CF.cCtaCod = CFE.cCtaCod " _
'    & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod " _
'    & " Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod " _
'    & " Inner Join RRHH    RH on RH.cPersCod = PP3.cPersCod " _
'    & " Inner Join Agencias Ag on Ag.cAgeCod = C.cAgeCodAct " _
'    & " Inner Join Constante K on K.nConsValor=C.cTpoCredCod and K.nConsCod = 3034 " _
'    & " WHERE PP.nPrdPersRelac = 20 and PP2.nPrdPersRelac=35 and PP3.nPrdPersRelac=28 " _
'    & " and CFE.nPrdEstado = Pro.nPrdEstado and CFE.nPrdEstado = 2091 " _
'    & " and CF.dEjecutado between convert (datetime , '" & Format(pdDesde, "mm/dd/yyyy") & "')" _
'    & " and convert (datetime , '" & Format(pdHasta, "mm/dd/yyyy") & "') "
'
'    If OptAge = True Then
'        lsSQL = lsSQL & " and C.cAgeCodAct ='" & psAgencia & "'"
'    Else
'        lsSQL = lsSQL & " and C.cAgeCodAct  in (" & psListaAgencias & " )"
'    End If
'    If Len(lcTipoM) > 0 Then
'        lsSQL = lsSQL & " and substring(C.cTpoCredCod,1,1) in (" & lcTipoM & ") "
'    End If
'
'    '******************
'   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
'      lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 1 & "' "
'   End If
'   If psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
'     lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 2 & "' "
'   End If

        
'On Error GoTo dError

   
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        'MsgBox " No Existen Datos para el reporte en la Agencia " & psAgencia, vbInformation, " Aviso"
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS DEVUELTAS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 152, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !Carta_Fianza & Space(1) & ofun.ImpreFormat(!Cliente, 30) & Space(2) _
                    & !Fecha_Emision & Space(2) & !Fecha_Emision & Space(2) & ofun.ImpreFormat(!Acreedor, 40) _
                    & Space(2) & !Analista
                
                If Not OptAge Then
                    lsCadImp = lsCadImp & Space(2) & !NombAgencia
                End If
                lsCadImp = lsCadImp & Space(3) & ofun.ImpreFormat(!Tip_Credito, 11, 0) 'madm 20101115
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS DEVUELTAS ", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 152, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
       
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo148005_CartasFianzaDevueltas = lsCadBuffer
    Set oFunI = Nothing
End Function


Public Function nRepo148001_CartasFianzaEmitidas(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdDesde As Date, ByVal pdHasta As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
        


Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim ofun As New COMFunciones.FCOMImpresion
Dim lsCadTpoCred As String
Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
   lsCadTpoCred = "1,2,3,4,5"
Else
    If psTipoCor = 1 Then
      lsCadTpoCred = lsCadTpoCred & ",'1'"
    End If
    If psTipoGraEmp = 1 Then
      lsCadTpoCred = lsCadTpoCred & ",'2'"
    End If
    If psTipoMedEmp = 1 Then
      lsCadTpoCred = lsCadTpoCred & ",'3'"
    End If
    If psTipoPeqEmp = 1 Then
      lsCadTpoCred = lsCadTpoCred & ",'4'"
    End If
    If psTipoMicEmp = 1 Then
      lsCadTpoCred = lsCadTpoCred & ",'5'"
    End If
    lsCadTpoCred = Mid(lsCadTpoCred, 2, Len(lsCadTpoCred) - 1)
End If
 '************ BRGO BASILEA II
lsSql = "Select Pro.cCtaCod as Carta_Fianza, CF.dAsignacion as Fecha_Emision, CF.dVencimiento as Fecha_Vencimiento," _
    & " P1.cPersNombre as Cliente, P2.cPersNombre as Acreedor, Pro.nSaldo as Importe," _
    & " RH.cUser as Analista,  " _
    & "   ( Select C.cConsDescripcion as Describe from Constante C where CFE.nPrdEstado=C.nConsValor and nConsCod=3001) as Estado, "
    lsSql = lsSql & " Ag.cAgeDescripcion as NombAgencia "
    lsSql = lsSql & " ,K.cConsDescripcion Tip_Credito " 'MADM TIPO CREDITO 20100928
    lsSql = lsSql & " From ColocCartaFianza CF  "
    lsSql = lsSql & " Inner Join Colocaciones C On C.cCtaCod = CF.cCtaCod "
    lsSql = lsSql & " Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod "
    lsSql = lsSql & " Inner Join ProductoPersona PP  on CF.cCtaCod=PP.cCtaCod "
    lsSql = lsSql & " Inner Join ProductoPersona PP2 on CF.cCtaCod=PP2.cCtaCod "
    lsSql = lsSql & " Inner Join ProductoPersona PP3 on CF.cCtaCod=PP3.cCtaCod "
    lsSql = lsSql & " Inner Join ColocCFEstado CFE on CF.cCtaCod = CFE.cCtaCod "
    lsSql = lsSql & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod "
    lsSql = lsSql & " Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod "
    lsSql = lsSql & " Inner Join RRHH    RH on RH.cPersCod = PP3.cPersCod "
    lsSql = lsSql & " Inner Join Agencias Ag on Ag.cAgeCod = C.cAgeCodAct"
    lsSql = lsSql & " Inner Join Constante K on K.nConsValor=C.cTpoCredCod and K.nConsCod = 3034" 'MADM TIPO CREDITO 20100928
    lsSql = lsSql & " WHERE PP.nPrdPersRelac = 20 and PP2.nPrdPersRelac=35 and PP3.nPrdPersRelac=28 "
    lsSql = lsSql & " and CFE.nPrdEstado = Pro.nPrdEstado "
    lsSql = lsSql & " and CF.dAsignacion between convert (datetime , '" & Format(pdDesde, "mm/dd/yyyy") & "')"
    lsSql = lsSql & " and convert (datetime , ' " & Format(pdHasta, "mm/dd/yyyy") & "') "
        
    If OptAge Then
        lsSql = lsSql & " and C.cAgeCodAct in (" & psAgencia & ")"
    Else
        lsSql = lsSql & " and C.cAgeCodAct in " & psListaAgencias
    End If

    If Len(lsCadTpoCred) > 0 Then
      lsSql = lsSql & " and substring(C.cTpoCredCod,1,1) In (" & lsCadTpoCred & ")"
    End If
    '******************
   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
      lsSql = lsSql & " and substring(CF.cCtaCod,9,1)='" & 1 & "' "
   End If
   If psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
     lsSql = lsSql & " and substring(CF.cCtaCod,9,1)='" & 2 & "' "
   End If
    
       
'On Error GoTo dError

   
    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psAgencia
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS EMITIDAS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 162, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
            Do While Not lrDataRep.EOF
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1
                
                'Cadena de Impresion
                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !Carta_Fianza & Space(1) & ofun.ImpreFormat(!Cliente, 30) & Space(2) _
                    & !Fecha_Emision & Space(2) & !Fecha_Vencimiento & Space(2) & ofun.ImpreFormat(!Acreedor, 40) _
                    & Space(2) & !Analista & Space(3) & ofun.ImpreFormat(!Estado, 12, 0)
                    
                If Not OptAge Then
                    lsCadImp = lsCadImp & Space(2) & !NombAgencia
                End If
                
                lsCadImp = lsCadImp & Space(1) & ofun.ImpreFormat(!Tip_Credito, 11, 0) 'madm 20101115
                
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS EMITIDAS", "DEL " & Format(pdDesde, "dd/mm/yyyy") & " - AL " & Format(pdHasta, "dd/mm/yyyy"), lnPage, 162, "", pnOpeCod)
                    lnLineas = 7
                End If
                .MoveNext
            Loop
        End With
        
       
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo148001_CartasFianzaEmitidas = lsCadBuffer
    Set oFunI = Nothing
End Function

'By Capi Acta 035-2007
Public Function nRepo148006_CartasFianzaCanceladas(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdCMACT As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psmensaje As String, Optional psImpresora As Impresoras = gEPSON, _
        Optional ByVal psDel As Date, Optional ByVal psAl As Date) As String
        
'*** PEAC - se agrego los parametros psDel,psAl
        
Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim ofun As New COMFunciones.FCOMImpresion

Dim P As String
Dim oFunI As New COMFunciones.FCOMVarImpresion
Dim lsAge, lsTipoC, lsMoneda As String

Dim nTotal As Double  'MAVM 05112009
Dim lcMone As String  'MAVM 05112009
Dim lcTipoM As String 'BRGO 20100615

oFunI.Inicia psImpresora

 
    If OptAge = True Then
        lsAge = psAgencia
    Else
        lsAge = Replace(Replace(psListaAgencias, "'", ""), " ", "")
    End If

'*** BRGO BASILEA II
    If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
        lcTipoM = "1,2,3,4,5"
    Else
        If psTipoCor = 1 Then
          lcTipoM = lcTipoM & ",1" ' Corporativo
        End If
        If psTipoGraEmp = 1 Then
          lcTipoM = lcTipoM & ",2" ' Grandes empresas
        End If
        If psTipoMedEmp = 1 Then
          lcTipoM = lcTipoM & ",3" ' Medianas empresas
        End If
        If psTipoPeqEmp = 1 Then
          lcTipoM = lcTipoM & ",4" ' Pequeñas empresas
        End If
        If psTipoMicEmp = 1 Then
          lcTipoM = lcTipoM & ",5" ' Microempresas
        End If
        lcTipoM = Mid(lcTipoM, 2, Len(lcTipoM) - 1)
    End If
'**********************

    If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
        lsMoneda = "1"
    ElseIf psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
        lsMoneda = "2"
    Else
        lsMoneda = "1,2"
    End If
   
    '***PEAC 20090615
    lsSql = " exec stp_sel_ObtieneCartaFianzaCanceladas '" & lsAge & "','" & lcTipoM & "','" & lsMoneda & "','" & Format(psDel, "yyyymmdd") & "','" & Format(psAl, "yyyymmdd") & "'"

'On Error GoTo dError

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing

    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia "
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "CANCELADAS DEL " & Format(psDel, "dd/mm/yyyy") & " AL " & Format(psAl, "dd/mm/yyyy"), lnPage, 178, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7

        With lrDataRep
            Do While Not lrDataRep.EOF
            
                lsCadImp = lsCadImp & "Moneda : " & IIf(!cCodMone = "1", "SOLES", "DOLAR") & oFunI.gPrnSaltoLinea
                lcMone = !cCodMone
                nTotal = 0
                
                Do While lrDataRep!cCodMone = lcMone
                lnIndice = lnIndice + 1
                lnLineas = lnLineas + 1

                'Cadena de Impresion
                'lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & ofun.ImpreFormat(!Carta_Fianza, 18) & Space(1) & "999" & Space(1) & ofun.ImpreFormat(!Cliente, 30) & Space(1)
                
                '*** PEAC 20090615 -  se agrego !Fecha_Cancel
                lsCadImp = lsCadImp & ofun.ImpreFormat(!Carta_Fianza, 18, 0) & Space(1) & ofun.ImpreFormat(!Num_Renovacion, 3, 0) & Space(1) & ofun.ImpreFormat(!Cliente, 30, 0) & Space(1) _
                    & !Fecha_Emision & Space(1) & !Fecha_Vencimiento & Space(1) & Format(!Fecha_Cancel, "dd/mm/yyyy") & Space(1) & Format(!Num_Poliza, "0000000") & Space(1) & ofun.ImpreFormat(!Importe, 12, 2, True) & Space(1) & ofun.ImpreFormat(!Acreedor, 30, 0) _
                    & Space(1) & ofun.ImpreFormat(!Modalidad, 12, 0)

                '*** PEAC 20090615
'                If Not OptAge Then
'                    lsCadImp = lsCadImp & Space(1) & ofun.ImpreFormat(!NombAgencia, 12, 0)
'                End If

                'MADM 20101115 - TIP CRED
                lsCadImp = lsCadImp & Space(1) & ofun.ImpreFormat(!NombAgencia, 12, 0) & Space(1) & ofun.ImpreFormat(!Tip_Credito, 11, 0)
                
                nTotal = nTotal + !Importe
                
                '*** FIN PEAC
                
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

                
                If lnIndice Mod 300 = 0 Then
                    lsCadBuffer = lsCadBuffer & lsCadImp
                    lsCadImp = ""
                End If
                
'                If lnLineas >= 55 Then
'                    lnPage = lnPage + 1
'                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
'                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
'                    lnLineas = 7
'                End If
                .MoveNext
                
                If lrDataRep.EOF Then
                    Exit Do
                End If
                
                If lnLineas >= 55 Then
                    lnPage = lnPage + 1
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 178, "", pnOpeCod)
                    lnLineas = 7
                End If
                
                Loop
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                lsCadImp = lsCadImp & "Total de Cartas Fianzas Vigentes: " & ofun.ImpreFormat(lnIndice, 5, 0)
                'MAVM 05112009 ***
                lsCadImp = lsCadImp & Space(53) & "Total: " & ofun.ImpreFormat(nTotal, 1, 2, True)
                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
              
         If lrDataRep.EOF Then
            Exit Do
         End If
                
                
            Loop
        End With
        
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
'        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'        lsCadImp = lsCadImp & "Total de Cartas Fianzas Vigentes: " & ofun.ImpreFormat(lnIndice, 5, 0)
'        'MAVM 05112009 ***
'        lsCadImp = lsCadImp & Space(53) & "Total: " & ofun.ImpreFormat(nTotal, 1, 2, True)
'
'        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo148006_CartasFianzaCanceladas = lsCadBuffer
Set oFunI = Nothing
End Function

'MADM 20111209 - 20101231
Public Function nRepo_CartasFianzaEmitidas(ByVal psTextoCarta As String, ByVal psTextoCarta1 As String, ByVal psTextoCarta2 As String, ByVal psCtaCod As String, ByVal pNROSEG As String, ByVal pFechas As String, _
        ByVal pSenores As String, ByVal pSOLICITANTE As String, ByVal pMODALIDAD As String, ByVal pFINALIDAD As String, ByVal pVENCIMIENTO As String, ByVal nSaldo As String, _
        ByVal pDIRECCION As String, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
        
Dim lsCadImp As String
Dim lsCartaModelo As String

Dim lsSaltoLin As String
Dim gPrnNegritaON As String
Dim gPrnNegritaOFF As String

Dim ofcar As New COMFunciones.FCOMCadenas
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
Dim pNRO As String
Dim cuenta As String
Dim cuenta1 As String
Dim cuenta2 As String
Dim cant As Integer

lsSaltoLin = Chr$(10)
lsCadImp = ""
nRepo_CartasFianzaEmitidas = ""
oFunI.Inicia psImpresora


pNRO = Left(psCtaCod, 3) & "-" & Mid(psCtaCod, 4, 2) & "-" & Mid(psCtaCod, 6, 3) & "-" & Mid(psCtaCod, 9, 10)

    If psCtaCod = "" Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psCtaCod
        Exit Function
    Else
        cant = Len(Trim(pFINALIDAD))
        cuenta = ofun.Justifica(pFINALIDAD, 90) 'ofun.ImpreFormat(ofun.QuiebreTexto(cFinalidad, 1), 85)
        
        If cant > 185 Then
            cuenta1 = Mid(Trim(pFINALIDAD), 91, 90)
            cuenta2 = Mid(Trim(pFINALIDAD), 181, cant - 181)
            lsCartaModelo = psTextoCarta
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<NRO>>", pNRO, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<NROSEG>>", pNROSEG, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<FECHA>>", pFechas, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<SEÑORES>>", pSenores, , 2, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<SOLICITANTE>>", pSOLICITANTE, , 2, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<MONTO>>", nSaldo, , 2, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<MODALIDAD>>", pMODALIDAD, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<FINALIDAD>>", cuenta, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<FINALIDAD1>>", cuenta1, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<FINALIDAD2>>", cuenta2, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<VENCIMIENTO>>", Format(pVENCIMIENTO, "dddd,d mmmm yyyy"), , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<DIRECCION>>", pDIRECCION, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCadImp = lsCadImp & lsCartaModelo & oFunI.gPrnSaltoPagina
        ElseIf cant >= 90 And cant <= 185 Then
            cuenta1 = Mid(Trim(pFINALIDAD), 91, cant)
            lsCartaModelo = psTextoCarta1
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<NRO>>", pNRO, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<NROSEG>>", pNROSEG, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<FECHA>>", pFechas, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<SEÑORES>>", pSenores, , 2, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<SOLICITANTE>>", pSOLICITANTE, , 2, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<MONTO>>", nSaldo, , 2, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<MODALIDAD>>", pMODALIDAD, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<FINALIDAD>>", cuenta, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<FINALIDAD1>>", cuenta1, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<VENCIMIENTO>>", Format(pVENCIMIENTO, "dddd,d mmmm yyyy"), , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<DIRECCION>>", pDIRECCION, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCadImp = lsCadImp & lsCartaModelo & oFunI.gPrnSaltoPagina
        Else
            lsCartaModelo = psTextoCarta2
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<NRO>>", pNRO, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<NROSEG>>", pNROSEG, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<FECHA>>", pFechas, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<SEÑORES>>", pSenores, , 2, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<SOLICITANTE>>", pSOLICITANTE, , 2, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<MONTO>>", nSaldo, , 2, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<MODALIDAD>>", pMODALIDAD, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<FINALIDAD>>", cuenta, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<VENCIMIENTO>>", Format(pVENCIMIENTO, "dddd,d mmmm yyyy"), , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCartaModelo = oFunI.gPrnNegritaON & Replace(lsCartaModelo, "<<DIRECCION>>", pDIRECCION, , 1, vbTextCompare) & oFunI.gPrnNegritaOFF
            lsCadImp = lsCadImp & lsCartaModelo & oFunI.gPrnSaltoPagina
        End If
    End If
    nRepo_CartasFianzaEmitidas = lsCadImp
    Set oFunI = Nothing
Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Emision Carta Fianza  <<nImprimeAvisoVencimiento>>", Err.Description
End Function


Public Function nRepo_CartasFianzaEmitidasPoliza(ByVal psTextoCarta As String, ByVal psTextoCarta1 As String, ByVal psTextoCarta2 As String, ByVal psCtaCod As String, ByVal dFecCrea As String, ByVal psFechaSis As Date, ByVal dFecVenA As Date, ByVal dFecVenN As String, _
        ByVal cfinalidad As String, ByVal cTitular As String, ByVal dAcreedor As String, ByVal nSaldo As String, _
        ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String
        
Dim lsCadImp As String
Dim lsCartaModelo As String

Dim lsSaltoLin As String
Dim gPrnNegritaON As String
Dim gPrnNegritaOFF As String

Dim ofcar As New COMFunciones.FCOMCadenas
Dim ofun As New COMFunciones.FCOMImpresion
Dim oFunI As New COMFunciones.FCOMVarImpresion
Dim cuenta As String
Dim cuenta1 As String
Dim cuenta2 As String
Dim cant As Integer

lsSaltoLin = Chr$(10)
lsCadImp = ""
nRepo_CartasFianzaEmitidasPoliza = ""
oFunI.Inicia psImpresora

    If psCtaCod = "" Then
        psmensaje = " No Existen Datos para el reporte en la Agencia " & psCtaCod
        Exit Function
    Else
            cant = Len(Trim(cfinalidad))
            cuenta = ofun.Justifica(cfinalidad, 90) 'ofun.ImpreFormat(ofun.QuiebreTexto(cFinalidad, 1), 85)
            
            If cant > 185 Then
                cuenta1 = Mid(Trim(cfinalidad), 91, 90)
                cuenta2 = Mid(Trim(cfinalidad), 181, cant - 181)
                lsCartaModelo = psTextoCarta2
                lsCartaModelo = Replace(lsCartaModelo, "<<dFecha>>", Format(psFechaSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<cTitular>>", Trim(cTitular), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<cCtaCod>>", psCtaCod, , 2, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<dFecCrea>>", dFecCrea, , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<cFinalidad>>", Trim(cuenta), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<cFinalidad1>>", Trim(cuenta1), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<cFinalidad2>>", Trim(cuenta2), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<nMonto>>", Trim(nSaldo), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<dFecVenA>>", Format(dFecVenA, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<dAcreedor>>", dAcreedor, , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<dFecVenN>>", dFecVenN, , 1, vbTextCompare)
                lsCadImp = lsCadImp & lsCartaModelo & oFunI.gPrnSaltoPagina
            ElseIf cant >= 90 And cant <= 185 Then
                cuenta1 = Mid(Trim(cfinalidad), 91, cant)
                lsCartaModelo = psTextoCarta1
                lsCartaModelo = Replace(lsCartaModelo, "<<dFecha>>", Format(psFechaSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<cTitular>>", Trim(cTitular), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<cCtaCod>>", psCtaCod, , 2, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<dFecCrea>>", dFecCrea, , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<cFinalidad>>", Trim(cuenta), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<cFinalidad1>>", Trim(cuenta1), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<nMonto>>", Trim(nSaldo), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<dFecVenA>>", Format(dFecVenA, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<dAcreedor>>", dAcreedor, , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<dFecVenN>>", dFecVenN, , 1, vbTextCompare)
                lsCadImp = lsCadImp & lsCartaModelo & oFunI.gPrnSaltoPagina
            Else
                lsCartaModelo = psTextoCarta
                lsCartaModelo = Replace(lsCartaModelo, "<<dFecha>>", Format(psFechaSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<cTitular>>", Trim(cTitular), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<cCtaCod>>", psCtaCod, , 2, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<dFecCrea>>", dFecCrea, , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<cFinalidad>>", Trim(cfinalidad), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<nMonto>>", Trim(nSaldo), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<dFecVenA>>", Format(dFecVenA, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<dAcreedor>>", dAcreedor, , 1, vbTextCompare)
                lsCartaModelo = Replace(lsCartaModelo, "<<dFecVenN>>", dFecVenN, , 1, vbTextCompare)
                lsCadImp = lsCadImp & lsCartaModelo & oFunI.gPrnSaltoPagina
            End If
    End If
    nRepo_CartasFianzaEmitidasPoliza = lsCadImp
    Set oFunI = Nothing
Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Emision Carta Fianza  <<nImprimeAvisoVencimiento>>", Err.Description
End Function
'loRs1.nRepo_CartasFianzaEmitidas(Format(CLng(txtNumPoliza.Text), "0000000"), lsFechas, PstaNombre(lblNomAcreedor, True), PstaNombre(lblNomcli, True), lblModalidad.Caption, lblFinalidad.Caption, lsFechas1, montox, cDirecAgencia, lsmensaje, gImpresora)
Public Function ImprimeCartillaEmision(ByVal psCtaCod As String, ByVal pNROSEG As String, ByVal psFechaSis As String, ByVal dFecVenN As String, _
        ByVal pSenores As String, ByVal cSolicitante As String, ByVal nSaldo As String, ByVal cModalidad As String, ByVal cfinalidad As String, _
        ByVal cDireccion As String, ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsCad As String
Dim nInteres As Double
Dim nMeses As Integer
Dim OiMO As New COMFunciones.FCOMVarImpresion
Dim ofun As New COMFunciones.FCOMImpresion
Dim i As Byte
Dim BON As String, BOFF As String

If psImpresora = gEPSON Then
    BON = Chr$(27) & Chr$(69)
    BOFF = Chr(27) & Chr$(70)
Else
    BOFF = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(48) & Chr$(66)
    BON = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(51) & Chr$(66)
End If

BON = PrnSet("B+")
BOFF = PrnSet("B-")
lsCad = ""
OiMO.Inicia psImpresora

'lsCad = lsCad & OiMO.gPrnCondensadaON
lsCad = lsCad & OiMO.gPrnSaltoLinea

    lsCad = lsCad & OiMO.gPrnSaltoLinea & OiMO.gPrnSaltoLinea & OiMO.gPrnSaltoLinea
    lsCad = lsCad & ofun.Centra("CARTA FIANZA N° ", 80) & pNROSEG & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(64) & Format(psFechaSis, "dddd,d mmmm yyyy") & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & BON & "Señores: " & BOFF
    lsCad = lsCad & Space(10) & BON & pSenores & BOFF & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    
    lsCad = lsCad & Space(10) & "Cuidad.- " & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    
    lsCad = lsCad & Space(10) & "Muy Señores Nuestros:"
    lsCad = lsCad & Space(14) & ofun.Justifica("A solicitud de " & cSolicitante & ", otorgamos por el presente ", 90) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("documento una fianza solidaria, irrevocable, incondicional, de ejecución inmediata, con renuncia", 90) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("expresa al beneficio de excusión e indivisible, a favor de ustedes garantizando a ", 90) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica(cSolicitante & " por la suma de " & Trim(nSaldo), 90) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("a fin de garantizar la " & cModalidad & " de la " & cfinalidad & " por la suma de " & Trim(nSaldo), 90) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    
    lsCad = lsCad & Space(14) & ofun.Justifica("Dejamos claramente establecido que la presente Carta Fianza no podrá ser usada", 100) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("para operaciones comprendidas en la prohibición indicada en el inciso “5” del Articulo 217 de ", 100) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("la Ley 26702, Ley General del Sistema Financiero y del Sistema de Seguros y Orgánica de la", 100) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("Superintendencia de Banca y Seguros.", 42) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & OiMO.gPrnSaltoLinea & OiMO.gPrnSaltoLinea
    
    lsCad = lsCad & Space(14) & ofun.Justifica("Por efecto de este compromiso la CAJA MUNICIPAL DE MAYNAS S.A asume con su fiado", 90) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("las responsabilidades en que éste llegara a incurrir siempre que el monto de las mismas no exceda", 100) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("por ningún motivo de la suma antes mencionada y que estén estrictamente vinculadas al cumplimiento", 100) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("de lo arriba indicado.", 22) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & OiMO.gPrnSaltoLinea & OiMO.gPrnSaltoLinea
        
    lsCad = lsCad & Space(14) & ofun.Justifica("La presente garantía rige a partir de la fecha y vencerá el " & dFecVenN, 90) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("Cualquier  reclamo en virtud de esta garantía deberá ceñirse estrictamente a lo estipulado por el", 100) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("Art. 1898 del Código Civil y deberá ser formulado por vía notarial y en nuestra oficina ubicada", 100) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("en " & Trim(cDireccion), 50) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    
    lsCad = lsCad & Space(30) & "Atentamente" & OiMO.gPrnSaltoLinea
    lsCad = lsCad & OiMO.gPrnCondensadaOFF
    ImprimeCartillaEmision = lsCad
End Function

Public Function ImprimeCartillaRenovacion(ByVal psCtaCod As String, ByVal pNROSEG As String, ByVal dFecCrea As String, ByVal psFechaSis As Date, ByVal dFecVenA As Date, ByVal dFecVenN As String, _
        ByVal cfinalidad As String, ByVal cTitular As String, ByVal dAcreedor As String, ByVal nSaldo As String, _
        ByRef psmensaje As String, Optional ByVal psImpresora As Impresoras = gEPSON) As String

Dim lsCad As String
Dim nInteres As Double
Dim nMeses As Integer
Dim OiMO As New COMFunciones.FCOMVarImpresion
Dim ofun As New COMFunciones.FCOMImpresion
Dim i As Byte
Dim BON As String, BOFF As String

If psImpresora = gEPSON Then
    BON = Chr$(27) & Chr$(69)
    BOFF = Chr(27) & Chr$(70)
Else
    BOFF = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(48) & Chr$(66)
    BON = Chr$(27) & Chr$(40) & Chr$(115) & Chr$(51) & Chr$(66)
End If

BON = PrnSet("B+")
BOFF = PrnSet("B-")
lsCad = ""
OiMO.Inicia psImpresora

'lsCad = lsCad & OiMO.gPrnCondensadaON
lsCad = lsCad & OiMO.gPrnSaltoLinea

    lsCad = lsCad & OiMO.gPrnSaltoLinea & OiMO.gPrnSaltoLinea & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(64) & Format(psFechaSis, "dddd,d mmmm yyyy") & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & BON & "Señores: " & BOFF & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & BON & ofun.Centra(cTitular, 90) & BOFF & OiMO.gPrnSaltoLinea
    'lsCad = lsCad & Space(30) & Trim(cTitular) & OiMO.gPrnNegritaON & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "Cuidad: " & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & ofun.Justifica("REF: Nuestra Carta Fianza Nº" & BON & pNROSEG & BOFF & "con credito " & BON & psCtaCod & BOFF & " del " & dFecCrea & "para la obra " & BON & Trim(cfinalidad) & BOFF & "por " & Trim(nSaldo) & " con vencimiento " & Format(dFecVenA, "dddd,d mmmm yyyy") & " a favor de ustedes y a/c. de ", 90) & OiMO.gPrnSaltoLinea
'    lsCad = lsCad & Space(10) & ofun.Justifica("para la obra " & Trim(cFinalidad), 90) & OiMO.gPrnSaltoLinea
'    lsCad = lsCad & Space(10) & ofun.Justifica("por " & Trim(nSaldo) & " con vencimiento " & Format(dFecVenA, "dddd,d mmmm yyyy") & "a favor de ustedes y a/c. de ", 90) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & BON & ofun.Centra(dAcreedor, 90) & BOFF & OiMO.gPrnSaltoLinea
    'lsCad = lsCad & Space(30) & Trim(dAcreedor) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "Estimados señores:" & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(14) & ofun.Justifica("Sírvase a tomar nota de que, a solicitud de nuestros garantizados, hemos ", 90) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & "Procedido a:" & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(14) & ofun.Justifica("Prorogar el plazo de vencimiento hasta el " & Format(dFecVenN, "dddd,d mmmm yyyy"), 90) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(14) & ofun.Justifica("Manteniéndose vigente todos los demás términos y consideraciones de las misma.", 90) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(14) & ofun.Justifica("Cualquier reclamo en virtud de esta garantía debe ceñirse a lo estipulado por el Art. 1898 del Código Civil y deberá ser formulado por la vía notarial en el horario de atención al público y en nuestra Oficina ubicada en Jr. Prospero Nº 791,Iquitos", 90) & OiMO.gPrnSaltoLinea
    lsCad = lsCad & OiMO.gPrnSaltoLinea & OiMO.gPrnSaltoLinea
    lsCad = lsCad & Space(10) & OiMO.gPrnSaltoLinea
    
    lsCad = lsCad & Space(30) & "Atentamente" & OiMO.gPrnSaltoLinea
    lsCad = lsCad & OiMO.gPrnCondensadaOFF
    ImprimeCartillaRenovacion = lsCad
End Function

Public Function PrnSet(Code As String, Optional nValor As Integer) As String
If nValor = 12 Or nValor = 10 Then
   nValor = nValor - 1
End If
Select Case Code
 Case "B+": PrnSet = Chr$(27) & Chr$(69) 'Bold On
 Case "B-": PrnSet = Chr$(27) & Chr$(70) 'Bold Off
 Case "U+": PrnSet = Chr$(27) & Chr$(45)  'Underline On
 Case "U-": PrnSet = Chr$(27) & Chr$(46) 'Underline Off
 Case "I+": PrnSet = Chr$(27) & Chr$(52) 'Italic On
 Case "I-": PrnSet = Chr$(27) & Chr$(53) 'Italic Off
 Case "W+": PrnSet = Chr$(27) & Chr$(87) 'Doble Ancho On
 Case "W-": PrnSet = Chr$(27) & Chr$(20) 'Doble Ancho Off
 Case "C+": PrnSet = Chr$(27) & Chr$(15) 'Condensado On
 Case "C-": PrnSet = Chr$(27) & Chr$(18) 'Condensado Off
 Case "Rm": PrnSet = Chr$(27) & Chr$(107) & Chr$(0) 'Roman
 Case "Ss": PrnSet = Chr$(27) & Chr$(107) & Chr$(1) 'Sans Serif
 Case "Co": PrnSet = Chr$(27) & Chr$(107) & Chr$(2) 'Courier
 Case "1.5": PrnSet = Chr$(27) & Chr$(48) ' 1 1/2 espacios
 Case "MI": PrnSet = Chr$(27) & Chr$(108) & Chr$(nValor) 'Margen Izquierdo
 Case "MD": PrnSet = Chr$(27) & Chr$(81) & Chr$(nValor)  'Margen Derecho
 Case "10CPI": PrnSet = Chr$(27) & Chr$(80)
 Case "12CPI": PrnSet = Chr$(27) & Chr$(77)
 Case "15CPI": PrnSet = Chr$(27) & Chr$(103)
 Case "EspN": PrnSet = Chr$(27) & Chr$(50)     'Espaciado Normal 4.5/72
 Case "Esp":  PrnSet = Chr$(27) & Chr$(65) & Chr$(nValor) 'Espaciado nValor/72 pulg.
End Select
End Function



'Public Function nRepo148002_CartasFianzaVigentes(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
'        ByVal pdCMACT As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
'        ByRef psMonedaE As Integer, ByRef psTipoC As Integer, ByRef psTipoM As Integer, _
'        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psmensaje As String, Optional psImpresora As Impresoras = gEPSON) As String
'
'Dim lsSQL As String
'Dim lrDataRep As New ADODB.Recordset
'Dim loDataRep As COMDColocPig.DCOMColPFunciones
'Dim lsCadImp As String
'Dim lsCadBuffer As String
'
'Dim lnIndice As Long
'Dim lnLineas As Integer
'Dim lnPage As Integer
'Dim lsOperaciones As String
'Dim ofun As New COMFunciones.FCOMImpresion
'
'Dim P As String
'
'Dim oFunI As New COMFunciones.FCOMVarImpresion
'oFunI.Inicia psImpresora
'
'
'lsSQL = "Select CF.cCtaCod as Carta_Fianza, CF.dAsignacion as Fecha_Emision, CF.dVencimiento as Fecha_Vencimiento," _
'    & " P1.cPersNombre as Cliente, P2.cPersNombre as Acreedor, CC.nMontoCol as Importe," _
'    & " RH.cUser as Analista, " _
'    & " Ag.cAgeDescripcion as NombAgencia, " _
'    & " ( Select C.cConsDescripcion  from Constante C where Pro.nPrdEstado=C.nConsValor and nConsCod=3001) as Estado " _
'    & " From Producto Pro " _
'    & " Inner Join colocaciones CC on CC.cCtaCod = Pro.cCtaCod " _
'    & " " _
'    & " Inner Join ColocCartaFianza CF on CF.cCtaCod = Pro.cCtaCod " _
'    & " Inner Join ProductoPersona PP  on CF.cCtaCod=PP.cCtaCod " _
'    & " Inner Join ProductoPersona PP2 on CF.cCtaCod=PP2.cCtaCod " _
'    & " Inner Join ProductoPersona PP3 on CF.cCtaCod=PP3.cCtaCod " _
'    & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod " _
'    & " Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod " _
'    & " Inner Join RRHH    RH on RH.cPersCod = PP3.cPersCod " _
'    & " Inner Join Agencias Ag on Ag.cAgeCod=Substring(CF.cCtaCod,4,2)" _
'    & " WHERE PP.nPrdPersRelac = 20 and PP2.nPrdPersRelac=35 and PP3.nPrdPersRelac=28 " _
'    & " And Pro.nPrdEstado=" & gColocEstVigNorm & "" _
'    & "  "
'
'    If OptAge = True Then
'        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)='" & psAgencia & "'"
'    Else
'        lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2) in (" & psListaAgencias & " )"
'    End If
'    'lsSQL = lsSQL & " and substring(CF.cCtaCod,4,2)=" & psAgencia
'
'
'   '************
'   If psTipoC = 0 And psTipoM = 1 Then 'tipo = Comercial
'       lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 2 & "' "
'   End If
'
'   If psTipoC = 1 And psTipoM = 0 Then 'tipo = MicroEmpresa
'      lsSQL = lsSQL & " and substring(CF.cCtaCod,6,1)='" & 1 & "' "
'   End If
'   '******************
'   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
'    lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 1 & "' "
'   End If
'   If psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
'    lsSQL = lsSQL & " and substring(CF.cCtaCod,9,1)='" & 2 & "' "
'   End If
'
''On Error GoTo dError
'
'    Set loDataRep = New COMDColocPig.DCOMColPFunciones
'        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSQL)
'    Set loDataRep = Nothing
'
'    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
'        psmensaje = " No Existen Datos para el reporte en la Agencia "
'        Exit Function
'    Else
'        lnLineas = 0
'        lnPage = 1
'
'        'CABECERA
'        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ",  "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
'
'        lnIndice = 0:  lnLineas = 7
'
'        With lrDataRep
'            Do While Not lrDataRep.EOF
'                lnIndice = lnIndice + 1
'                lnLineas = lnLineas + 1
'
'                'Cadena de Impresion
'                lsCadImp = lsCadImp & ofun.ImpreFormat(lnIndice, 5, 0) & Space(1) & !Carta_Fianza & Space(1) & ofun.ImpreFormat(!Cliente, 30) & Space(2) _
'                    & !Fecha_Emision & Space(2) & !Fecha_Vencimiento & Space(2) & ofun.ImpreFormat(!Acreedor, 40) _
'                    & Space(2) & !Analista
'
'                If Not OptAge Then
'                    lsCadImp = lsCadImp & Space(2) & !NombAgencia
'                End If
'                lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
'
'
'                If lnIndice Mod 300 = 0 Then
'                    lsCadBuffer = lsCadBuffer & lsCadImp
'                    lsCadImp = ""
'                End If
'
'                If lnLineas >= 55 Then
'                    lnPage = lnPage + 1
'                    lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
'                    lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
'                    lnLineas = 7
'                End If
'                .MoveNext
'            Loop
'        End With
'
'
'        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
'
'    End If
'    nRepo148002_CartasFianzaVigentes = lsCadBuffer
'Set oFunI = Nothing
'End Function

'madm 20111227
Public Function nRepo148007_CartasFianzanuevas(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdCMACT As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psmensaje As String, Optional psImpresora As Impresoras = gEPSON, Optional ByVal psDel As Date, Optional ByVal psAl As Date) As String
        
Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim ofun As New COMFunciones.FCOMImpresion

Dim P As String, sAges As String, sTipo As String, sMone As String
Dim lcAge As String, lcMone As String, lnTot As Double

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora

    If OptAge = True Then
        sAges = 1
    Else
        sAges = 2
    End If
    
  
   If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
        sTipo = "1,2,3,4,5"
    Else
        If psTipoCor = 1 Then
          sTipo = sTipo & ",1"
        End If
        If psTipoGraEmp = 1 Then
          sTipo = sTipo & ",2"
        End If
        If psTipoMedEmp = 1 Then
          sTipo = sTipo & ",3"
        End If
        If psTipoPeqEmp = 1 Then
          sTipo = sTipo & ",4"
        End If
        If psTipoMicEmp = 1 Then
          sTipo = sTipo & ",5"
        End If
        sTipo = Mid(sTipo, 2, Len(sTipo) - 1)
    End If
'**********************
   


   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
    sMone = "1"
   ElseIf psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
    sMone = "2"
    Else
    sMone = "1,2"
   End If
        
    lsSql = " exec stp_sel_ReporteCartasFianzasNuevasExcel '" & sMone & "','" & sTipo & "','" & sAges & "','" & Format(psDel, "yyyymmdd") & "','" & Format(psAl, "yyyymmdd") & "'"

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia "
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1

        'CABECERA
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "NUEVAS AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 179, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        
        With lrDataRep
        
            
            
            Do While Not lrDataRep.EOF
            
                lsCadImp = lsCadImp & "Agencia : " & !NombAgencia & oFunI.gPrnSaltoLinea
                lcAge = !cCodAge
                                
                Do While lrDataRep!cCodAge = lcAge
                    lsCadImp = lsCadImp & "Moneda : " & IIf(!cCodMone = "1", "SOLES", "DOLAR") & oFunI.gPrnSaltoLinea
                    lcMone = !cCodMone
                                
                    lnTot = 0
                    Do While lrDataRep!cCodMone = lcMone And lrDataRep!cCodAge = lcAge
            
                        lnIndice = lnIndice + 1
                        lnLineas = lnLineas + 1

                        lsCadImp = lsCadImp & ofun.ImpreFormat(!Carta_Fianza, 18, 0) & Space(1) & ofun.ImpreFormat(!Num_Renovacion, 3, 0) & Space(1) & ofun.ImpreFormat(!Cliente, 30, 0) & Space(1) & ofun.ImpreFormat(!TpoCred, 24, 0) & Space(1) _
                            & !Fecha_Emision & Space(1) & !Fecha_Vencimiento & Space(1) & Format(!Num_Poliza, "0000000") & Space(1) & ofun.ImpreFormat(!Importe, 12, 2, True) & Space(1) & ofun.ImpreFormat(!Acreedor, 30, 0) _
                            & Space(1) & ofun.ImpreFormat(!Modalidad, 12, 0)
                            
                            If Not OptAge Then
                                lsCadImp = lsCadImp & Space(1) & ofun.ImpreFormat(!NombAgencia, 12, 0)
                            End If
                            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                
                            If lnIndice Mod 300 = 0 Then
                                lsCadBuffer = lsCadBuffer & lsCadImp
                                lsCadImp = ""
                            End If
                
                            lnTot = lnTot + !Importe
                
                            lrDataRep.MoveNext
                            
                            If lrDataRep.EOF Then
                                Exit Do
                            End If
                
                            If lnLineas >= 55 Then
                                lnPage = lnPage + 1
                                lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                                lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "NUEVAS AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
                                lnLineas = 7
                            End If
                
                    Loop
                    
                    lsCadImp = lsCadImp & String(147, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & "Totales :" & Space(75) & ofun.ImpreFormat(lnTot, 12, 2, True)
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    
                    If lrDataRep.EOF Then
                        Exit Do
                    End If
                Loop
                If lrDataRep.EOF Then
                    Exit Do
                End If
            Loop
            
        End With
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Total de Cartas Fianzas Nuevas: " & ofun.ImpreFormat(lnIndice, 5, 0)
       
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
        
    End If
    nRepo148007_CartasFianzanuevas = lsCadBuffer
Set oFunI = Nothing
End Function

'WIOR 20120404
Public Function EmisionCF(ByVal pnCodCta As String, ByVal pbAvalado As Boolean, ByRef psmensaje As String, ByVal pdFecha As Date, ByVal psDireccion As String) As String

Dim oCF As COMDCartaFianza.DCOMCartaFianza
Dim rsCF As ADODB.Recordset

Dim dfechafin As Date
Dim sTitular As String
Dim sAval As String
Dim sFechaVen As String
Dim sFechaAct As String
Dim sMonto As String


Dim lsCadImp As String
Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lnNegritaON As String
Dim lnNegritaOFF As String

Dim ofun As New COMFunciones.FCOMImpresion
Dim ofcar As New COMFunciones.FCOMCadenas
Dim oFunI As New COMFunciones.FCOMVarImpresion

oFunI.Inicia gEPSON

    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr(27) & Chr$(70)
    
       
    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set rsCF = oCF.RecuperaCartaFianzaDetalle(pnCodCta)
    
    lsCadImp = ""
    If rsCF Is Nothing Or (rsCF.BOF And rsCF.EOF) Then
        psmensaje = " No Existen Datos para el reporte"
        Exit Function
    Else
        dfechafin = CDate(rsCF!dVencimiento)
        sTitular = ofcar.PstaNombre(rsCF!cPersNombre, True)
        sAval = ofcar.PstaNombre(IIf(IsNull(rsCF!cAvalNombre), "", rsCF!cAvalNombre), True)
        sFechaVen = Format(dfechafin, "dd") & " de " & Format(dfechafin, "mmmm") & " del " & Format(dfechafin, "yyyy")
        sFechaAct = Format(pdFecha, "dd") & " de " & Format(pdFecha, "mmmm") & " del " & Format(pdFecha, "yyyy")
        sMonto = IIf(Mid(pnCodCta, 9, 1) = "1", "S/. ", "$. ") & Format(rsCF!nSaldo, "#,###0.00") & " " & "(" & UCase(ofcar.NumLet(rsCF!nSaldo)) & IIf(Mid(pnCodCta, 9, 1) = "2", "", " Y " & IIf(InStr(1, rsCF!nSaldo, ".") = 0, "00", Mid(rsCF!nSaldo, InStr(1, rsCF!nSaldo, ".") + 1, 2)) & "/100 ") & IIf(Mid(pnCodCta, 9, 1) = "1", "NUEVOS SOLES)", "US DOLARES)")
        
        'Inicializacion de Líneas, Items y Páginas
        lnLineas = 1
        
        lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0) 'Tipo de Letra Sans Serif
        lsCadImp = lsCadImp & Chr$(27) & Chr$(108) & Chr$(30) 'MARGEN IZQUIERDO
        lsCadImp = lsCadImp & Chr$(27) & Chr$(81) & Chr$(30) 'MARGEN DERECHO
                  
        'Numero cuenta
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(75) & pnCodCta & oFunI.gPrnSaltoLinea

        
        'Numero Folio
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & ofun.Centra("CARTA FIANZA N° " & Format(rsCF!nPoliza, "0000000"), 95) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

        'Fecha
        lsCadImp = lsCadImp & Space(72) & sFechaAct & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        'Señores
        lsCadImp = lsCadImp & "Señores:" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & rsCF!cPersNomAcre & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        
        'Ciudad
        lsCadImp = lsCadImp & "Ciudad.-" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "--------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

        
        'Muy Señores
        lsCadImp = lsCadImp & "Muy Señores Nuestros:" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    
        
        'Parrafo 1
        Dim sParafo1 As String
        sParafo1 = Space(1) & "A solicitud de " & sTitular & ", otorgamos por el presente documento una" & _
                " fianza solidaria, irrevocable, incondicional, de ejecución inmediata, con renuncia expresa al" & _
                " beneficio de excusión e indivisible, a favor de ustedes" & IIf(pbAvalado = True, " garantizando a " & sAval, "") & "," & _
                " hasta por la suma de " & sMonto & ", a fin de garantizar" & _
                " la Carta Fianza por " & rsCF!Modalidad & ", referente a la: " & Chr$(34) & rsCF!cfinalidad & Chr$(34) & "."
        sParafo1 = ofun.JustificaTextoSinCarEsp(sParafo1, 95)
        lsCadImp = lsCadImp & sParafo1 & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

        'Parrafo 2
         Dim sParafo2 As String
         sParafo2 = Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & _
                "Dejamos claramente establecido que la presente Carta Fianza no podrá ser usada" & _
                " para operaciones comprendidas en la prohibición indicada en el inciso " & Chr$(34) & "5" & Chr$(34) & " del Articulo 217 de la" & _
                " Ley 26702, Ley General del Sistema Financiero y del Sistema de Seguros y Orgánica de la" & _
                " Superintendencia de Banca y Seguros."
        sParafo2 = ofun.JustificaTextoSinCarEsp(sParafo2, 95)
        lsCadImp = lsCadImp & sParafo2 & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        
        
        'Parrafo 3
        Dim sParafo3 As String
        sParafo3 = Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & _
                "Por efecto de este compromiso la CAJA MUNICIPAL DE MAYNAS S.A asume con su fiado" & _
                " las responsabilidades en que éste llegara a incurrir siempre que el monto de las mismas" & _
                " no exceda por ningún motivo de la suma antes mencionada y que estén estrictamente vinculadas" & _
                " cumplimiento de lo arriba indicado."
        sParafo3 = ofun.JustificaTextoSinCarEsp(sParafo3, 95)
        lsCadImp = lsCadImp & sParafo3 & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        
        
        
        'Parrafo 4
        Dim sParafo4 As String
        sParafo4 = Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & Chr$(32) & _
                "La presente garantía rige a partir de la fecha y vencerá el " & sFechaVen & "." & _
                " Cualquier  reclamo en virtud de esta garantía deberá ceñirse estrictamente a lo estipulado por el" & _
                " Art. 1898 del Código Civil y deberá ser formulado por vía notarial y en nuestra oficina ubicada en" & _
                " " & psDireccion & "."
        sParafo4 = ofun.JustificaTextoSinCarEsp(sParafo4, 95)
        lsCadImp = lsCadImp & sParafo4 & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        
        
        'final
        lsCadImp = lsCadImp & ofun.Centra("Atentamente,", 95) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & ofun.Centra("CAJA MUNICIPAL DE MAYNAS S.A", 95) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

    End If
    EmisionCF = lnNegritaON & lsCadImp & lnNegritaOFF
     
End Function
Public Function EmisionCFRenova(ByVal pnCodCta As String, ByVal pbAvalado As Boolean, ByRef psmensaje As String, ByVal pdFecha As Date, ByVal psDireccion As String, ByVal psAgencia As String) As String

    Dim oCF As COMDCartaFianza.DCOMCartaFianza
    Dim rsCF As ADODB.Recordset
    
    Dim dfechafin As Date
    Dim sTitular As String
    Dim sAval As String
    Dim sFechaVen As String
    Dim sFechaAct As String
    Dim sMonto As String
    
    
    Dim lsCadImp As String
    Dim lnIndice As Long
    Dim lnLineas As Integer
    Dim lnPage As Integer
    Dim lnNegritaON As String
    Dim lnNegritaOFF As String
    
    Dim ofun As New COMFunciones.FCOMImpresion
    Dim ofcar As New COMFunciones.FCOMCadenas
    Dim oFunI As New COMFunciones.FCOMVarImpresion
    Dim loRs As COMNCartaFianza.NCOMCartaFianzaValida


    Set loRs = New COMNCartaFianza.NCOMCartaFianzaValida
    Dim lrDataCF As ADODB.Recordset
    Set lrDataCF = loRs.RecuperaDatosGeneralesCF(pnCodCta)
    
    
    Dim dfechaini As Date
    Dim dfechafin02 As Date
    Dim lsFechas As String
    lsFechas = Format(lrDataCF!F_Asignacion, "dd") & " de " & Format(lrDataCF!F_Asignacion, "mmmm") & " del " & Format(lrDataCF!F_Asignacion, "yyyy")
    dfechaini = lrDataCF!dPrdEstado
    dfechafin02 = CDate(lrDataCF!dVenc)
    
    oFunI.Inicia gEPSON


    lnNegritaON = Chr$(27) & Chr$(69)
    lnNegritaOFF = Chr(27) & Chr$(70)
    
       
    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set rsCF = oCF.RecuperaCartaFianzaDetalle(pnCodCta)
    
    lsCadImp = ""
    If rsCF Is Nothing Or (rsCF.BOF And rsCF.EOF) Then
        psmensaje = " No Existen Datos para el reporte"
        Exit Function
    Else
        dfechafin = CDate(rsCF!dVencimiento)
        sTitular = ofcar.PstaNombre(rsCF!cPersNombre, True)
        sAval = ofcar.PstaNombre(IIf(IsNull(rsCF!cAvalNombre), "", rsCF!cAvalNombre))
        sFechaVen = Format(dfechafin, "dd/mm/yyyy")
        sFechaAct = Format(pdFecha, "dd") & " de " & Format(pdFecha, "mmmm") & " del " & Format(pdFecha, "yyyy")
        sMonto = IIf(Mid(pnCodCta, 9, 1) = "1", "S/.", "$.") & Format(rsCF!nSaldo, "#,###0.00")
        
        'Inicializacion de Líneas, Items y Páginas
        lnLineas = 1
        

        lsCadImp = lsCadImp & Chr$(27) + Chr$(107) + Chr$(0) 'Tipo de Letra Sans Serif
        lsCadImp = lsCadImp & Chr$(27) & Chr$(108) & Chr$(30) 'MARGEN IZQUIERDO
        lsCadImp = lsCadImp & Chr$(27) & Chr$(81) & Chr$(30) 'MARGEN DERECHO
                  
        'Numero Renovacion
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & Space(75) & "Renovación Nº " & rsCF!nRenovacion & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        
        'Fecha
        lsCadImp = lsCadImp & Space(60) & psAgencia & ", " & lsFechas & oFunI.gPrnSaltoLinea
        
        'Señores
        lsCadImp = lsCadImp & "Señores:" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & ofun.Centra(rsCF!cPersNomAcre, 95) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        
        'Ciudad
        lsCadImp = lsCadImp & "Ciudad.-" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "--------" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

        
        'Parrafo 1
        Dim sParafo1 As String
        If Not pbAvalado Then
            sParafo1 = Space(1) & "REF: Nuestra Carta Fianza Nº  " & Format(rsCF!nPoliza, "0000000") & " con crédito Nº " & pnCodCta & " del " & Format(dfechaini, "DD/MM/YYYY") & _
                    " para la obra " & Chr$(34) & rsCF!cfinalidad & Chr$(34) & _
                    " por " & sMonto & _
                    " con vencimiento  " & Format(dfechafin02, "DD/MM/YYYY") & " a favor de ustedes y a/c. de:"
                    
            sParafo1 = ofun.JustificaTextoSinCarEsp(sParafo1, 95)
            lsCadImp = lsCadImp & sParafo1 & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    
            lsCadImp = lsCadImp & ofun.Centra(sTitular, 95) & oFunI.gPrnSaltoLinea
        Else
            sParafo1 = Space(1) & "REF: Nuestra Carta Fianza Nº  " & Format(rsCF!nPoliza, "0000000") & " con crédito Nº " & pnCodCta & " del " & Format(dfechaini, "DD/MM/YYYY") & _
                    " para la obra " & Chr$(34) & rsCF!cfinalidad & Chr$(34) & _
                    " por " & sMonto & _
                    " con vencimiento  " & Format(dfechafin02, "DD/MM/YYYY") & " a favor de ustedes y a/c. de " & _
                    sTitular & ", garantizando a: "
                    
            sParafo1 = ofun.JustificaTextoSinCarEsp(sParafo1, 95)
            lsCadImp = lsCadImp & sParafo1 & oFunI.gPrnSaltoLinea
            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
    
            lsCadImp = lsCadImp & ofun.Centra(sAval, 95) & oFunI.gPrnSaltoLinea
        End If
        
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Estimados señores:" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Sírvase a tomar nota de que, a solicitud de nuestros garantizados, hemos" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Procedido a:" & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Prorrogar el plazo de vencimiento hasta el " & sFechaVen & "." & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & ofun.JustificaTextoSinCarEsp("Manteniéndose vigente todos los demás términos y consideraciones de las misma.", 95) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & ofun.JustificaTextoSinCarEsp("Cualquier reclamo en virtud de esta garantía deberá ceñirse a lo estipulado por el Art. 1898 del Código Civil y deberá ser formulado por vía notarial en el horario de atención al público y en nuestra Oficina Ubicada en " & psDireccion & ".", 95) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

        'Final
        lsCadImp = lsCadImp & ofun.Centra("Atentamente,", 95) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & ofun.Centra("CAJA MUNICIPAL DE MAYNAS S.A", 95) & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea

    End If
    EmisionCFRenova = lnNegritaON & lsCadImp & lnNegritaOFF
     
End Function
'WIOR -FIN


'******************RECO*20130715****
Public Function nRepo148008_CartasFianzaConsolidado(ByVal pnOpeCod As Long, ByVal psAgencia As String, _
        ByVal pdCMACT As Date, ByVal psUsuarioRep As String, ByRef psMonedaN As Integer, _
        ByRef psMonedaE As Integer, ByRef psTipoCor As Integer, ByRef psTipoGraEmp As Integer, _
        ByRef psTipoMedEmp As Integer, ByRef psTipoPeqEmp As Integer, ByRef psTipoMicEmp As Integer, _
        ByRef OptAge As Boolean, Optional psListaAgencias As String, Optional ByRef psmensaje As String, Optional psImpresora As Impresoras = gEPSON, _
        Optional ByRef psFecIni As String, Optional ByRef psFecFin As String) As String
        
Dim lsSql As String
Dim lrDataRep As New ADODB.Recordset
Dim loDataRep As COMDColocPig.DCOMColPFunciones
Dim lsCadImp As String
Dim lsCadBuffer As String

Dim lnIndice As Long
Dim lnLineas As Integer
Dim lnPage As Integer
Dim lsOperaciones As String
Dim ofun As New COMFunciones.FCOMImpresion

Dim P As String, sAges As String, sTipo As String, sMone As String
Dim lcAge As String, lcMone As String, lnTot As Double

Dim oFunI As New COMFunciones.FCOMVarImpresion
oFunI.Inicia psImpresora
    If OptAge = True Then
        sAges = psAgencia
    Else
        sAges = Replace(Replace(psListaAgencias, "'", ""), " ", "")
    End If
    If psTipoCor = 0 And psTipoGraEmp = 0 And psTipoMedEmp = 0 And psTipoPeqEmp = 0 And psTipoMicEmp = 0 Then
        sTipo = "1,2,3,4,5"
    Else
        If psTipoCor = 1 Then
          sTipo = sTipo & ",1"
        End If
        If psTipoGraEmp = 1 Then
          sTipo = sTipo & ",2"
        End If
        If psTipoMedEmp = 1 Then
          sTipo = sTipo & ",3"
        End If
        If psTipoPeqEmp = 1 Then
          sTipo = sTipo & ",4"
        End If
        If psTipoMicEmp = 1 Then
          sTipo = sTipo & ",5"
        End If
        sTipo = Mid(sTipo, 2, Len(sTipo) - 1)
    End If
   If psMonedaN = 1 And psMonedaE = 0 Then 'Moneda nacinal 1
    sMone = "1"
   ElseIf psMonedaN = 0 And psMonedaE = 1 Then ' Moneda extranjera 2
    sMone = "2"
    Else
    sMone = "1,2"
   End If
   '**************************
    lsSql = " exec stp_sel_ReporteCartasFianzasConsolidada '" & sAges & "','" & sMone & "','" & sTipo & "'" & ",'" & psFecIni & "','" & psFecFin & "'"

    Set loDataRep = New COMDColocPig.DCOMColPFunciones
        Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    Set loDataRep = Nothing
    
    If lrDataRep Is Nothing Or (lrDataRep.BOF And lrDataRep.EOF) Then
        psmensaje = " No Existen Datos para el reporte en la Agencia "
        Exit Function
    Else
        lnLineas = 0
        lnPage = 1
        lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 179, "", pnOpeCod)

        lnIndice = 0:  lnLineas = 7
        With lrDataRep
            Do While Not lrDataRep.EOF
            
                lsCadImp = lsCadImp & "Agencia : " & !NombAgencia & oFunI.gPrnSaltoLinea
                lcAge = !cCodAge
                                
                Do While lrDataRep!cCodAge = lcAge
                    lsCadImp = lsCadImp & "Moneda : " & IIf(!cCodMone = "1", "SOLES", "DOLAR") & oFunI.gPrnSaltoLinea
                    lcMone = !cCodMone
                                
                    lnTot = 0
                    Do While lrDataRep!cCodMone = lcMone And lrDataRep!cCodAge = lcAge
            
                        lnIndice = lnIndice + 1
                        lnLineas = lnLineas + 1

                        lsCadImp = lsCadImp & ofun.ImpreFormat(!Carta_Fianza, 18, 0) & Space(1) & ofun.ImpreFormat(!Num_Renovacion, 3, 0) & Space(1) & ofun.ImpreFormat(!Cliente, 30, 0) & Space(1) & ofun.ImpreFormat(!TpoCred, 24, 0) & Space(1) _
                            & !Fecha_Emision & Space(1) & !Fecha_Vencimiento & Space(1) & Format(!Num_Poliza, "0000000") & Space(1) & ofun.ImpreFormat(!Importe, 12, 2, True) & Space(1) & ofun.ImpreFormat(!Acreedor, 30, 0) _
                            & Space(1) & ofun.ImpreFormat(!Modalidad, 12, 0)
                            
                            If Not OptAge Then
                                lsCadImp = lsCadImp & Space(1) & ofun.ImpreFormat(!NombAgencia, 12, 0)
                            End If
                            lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
                
                            If lnIndice Mod 300 = 0 Then
                                lsCadBuffer = lsCadBuffer & lsCadImp
                                lsCadImp = ""
                            End If
                            lnTot = lnTot + !Importe
                            lrDataRep.MoveNext
                            If lrDataRep.EOF Then
                                Exit Do
                            End If
                            If lnLineas >= 55 Then
                                lnPage = lnPage + 1
                                lsCadImp = lsCadImp & oFunI.gPrnSaltoPagina
                                lsCadImp = lsCadImp & nRepoCabecera("LISTADO DE CARTAS FIANZAS ", "VIGENTES AL " & Format(pdCMACT, "dd/mm/yyyy"), lnPage, 150, "", pnOpeCod)
                                lnLineas = 7
                            End If
                    Loop
                    lsCadImp = lsCadImp & String(147, "-") & oFunI.gPrnSaltoLinea
                    lsCadImp = lsCadImp & "Totales :" & Space(75) & ofun.ImpreFormat(lnTot, 12, 2, True)
                    lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea & oFunI.gPrnSaltoLinea
                    If lrDataRep.EOF Then
                        Exit Do
                    End If
                Loop
                If lrDataRep.EOF Then
                    Exit Do
                End If
            Loop
        End With
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & oFunI.gPrnSaltoLinea
        lsCadImp = lsCadImp & "Total de Cartas Fianzas Vigentes: " & ofun.ImpreFormat(lnIndice, 5, 0)
        lsCadBuffer = lsCadBuffer & lsCadImp & oFunI.gPrnSaltoPagina
    End If
    nRepo148008_CartasFianzaConsolidado = lsCadBuffer
Set oFunI = Nothing
End Function
'*******************FIN RECO************************
'RECO20160120 ERS001-2016***********************************************
Public Function nRepo148009_CartasAvisoVencimineto(ByVal psAgencia As String, ByRef psMoneda As String, Optional psListaAgencias As String, _
        Optional ByVal psTipoCred As String, Optional ByVal psFecIni As String, Optional ByVal psFecFin As String) As ADODB.Recordset
        
    Dim loDataRep As New COMDColocPig.DCOMColPFunciones
    Dim lrDataRep As New ADODB.Recordset
    Dim lsSql As String
    Dim sAges As String
    
    If Replace(Replace(psListaAgencias, "'", ""), " ", "") = "" Then
        sAges = psAgencia
    Else
        sAges = Replace(Replace(psListaAgencias, "'", ""), " ", "")
        sAges = Replace(sAges, "(", "")
        sAges = Replace(sAges, ")", "")
    End If
   
    lsSql = " exec stp_sel_ReporteCartasFianzasPorVencerExcel '" & sAges & "','" & psMoneda & "','" & psTipoCred & "'" & ",'" & psFecIni & "','" & psFecFin & "'"
    Set nRepo148009_CartasAvisoVencimineto = loDataRep.dObtieneRecordSet(lsSql)
End Function
Public Function RegistraCartaAvisoVencimiento(ByVal psCtaCod As String, ByVal psUser As String) As String
    Dim loDataRep As New COMDColocPig.DCOMColPFunciones
    Dim lrDataRep As New ADODB.Recordset
    Dim lsSql As String
    Dim sAges As String
   
    lsSql = " exec stp_ins_CF_RegistraCartaVenc '" & psCtaCod & "','" & psUser & "'"
    Set lrDataRep = loDataRep.dObtieneRecordSet(lsSql)
    
    If Not (lrDataRep.EOF And lrDataRep.BOF) Then
        RegistraCartaAvisoVencimiento = lrDataRep!cNroCarta
    End If
    Set loDataRep = Nothing
End Function
'RECO FIN*************************************************************

