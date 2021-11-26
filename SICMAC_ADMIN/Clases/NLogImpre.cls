VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NLogImpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim pCorLin, pCorPag


Public Enum TpoImp
    TpoImpReqAdqui = 1
    TpoImpProveedor = 2
    TpoImpBS = 3
    TpoImpLogReq = 4
    TpoImpLogPrecios = 5
    TpoImpLogEstados = 6
End Enum

Private Function Cabecera(ByVal psIMP As TpoImp, ByVal psAge As String, ByVal pdAct As Date, _
ByVal pnPag As Integer, Optional ByVal pbPagCorta As Boolean = True, _
Optional ByVal psAdqNro As String = "") As String
    Dim clsDGnral As DLogGeneral
    Dim sCabe As String
    Dim sTit As String, sTitSub As String
    Dim sArea As String
    Dim nNroCol As Integer
    Set clsDGnral = New DLogGeneral
    
    sTit = "": sTitSub = ""
    Select Case psIMP
        Case TpoImpReqAdqui
            sTit = "L I S T A D O   D E   A D Q U I S I C I O N E S"
            sTitSub = "Adquisición : " & clsDGnral.GetsMovNro(psAdqNro)
        Case TpoImpProveedor
            sTit = "L I S T A D O   D E   P R O V E E D O R E S"
            sTitSub = ""
        Case TpoImpBS
            sTit = "L I S T A D O   D E   B I E N E S   Y   S E R V I C I O S"
            sTitSub = ""
    End Select
    sArea = "MODULO LOGISTICO"
    nNroCol = IIf(pbPagCorta = True, 135, 159)
    'Centra Título
    sTit = String(Round((nNroCol - Len(Trim(sTit))) / 2) - 25, " ") & sTit & String(Round((nNroCol - Len(Trim(sTit))) / 2) - 25, " ")
    'Centra SubTítulo
    sTitSub = String(Round((nNroCol - Len(Trim(sTitSub))) / 2) - 25, " ") & sTitSub & String(Round((nNroCol - Len(Trim(sTitSub))) / 2) - 25, " ")
    
    sCabe = sCabe & pCorLin
    sCabe = sCabe & Space(1) & ImpreFormat(UCase(psAge), 25, 0) & sTit & Space(11) & "PAGINA: " & Format(pnPag, "@@@@") & pCorLin
    sCabe = sCabe & Space(1) & ImpreFormat(sArea, 25, 0) & sTitSub & Space(4) & Format(pdAct & " " & Time, gsFormatoFechaHoraView) & pCorLin
    
    sCabe = sCabe & String(nNroCol, "-") & pCorLin
    Select Case psIMP
        Case TpoImpReqAdqui
            sCabe = sCabe & Space(2) & " ITEM      CODIGO                BIEN/SERVICIO           UNIDAD    CANTIDAD      PRECIO     SUBTOTAL" & pCorLin
        Case TpoImpProveedor
            'INICIO ORCR *****************
            'sCabe = sCabe & Space(2) & " ITEM             NOMBRE                         DIRECCION                 RUC              RUS            ESTADO" & pCorLin
            sCabe = sCabe & Space(2) & " ITEM             NOMBRE                         DIRECCION                 RUC              RUC              DOI            ESTADO" & pCorLin
            'FIN ORCR *****************
        Case TpoImpBS
            sCabe = sCabe & Space(2) & " ITEM       CODIGO                             NOMBRE                                      UNIDAD   " & pCorLin
    End Select
    sCabe = sCabe & String(nNroCol, "-") & pCorLin
    
    Cabecera = sCabe
    Set clsDGnral = Nothing
End Function


'Imp. Adquisiciones
Public Function ImpReqAdqui(ByVal psAgencia As String, ByVal pdactual As Date, _
ByVal poAdqNro As ADODB.Recordset) As String
    'Dim clsDAdq As DLogAdquisi
    Dim clsDReq As DLogRequeri
    Dim rs As ADODB.Recordset
    Dim sAdqNro As String
    Dim sImp As String
    Dim nLin As Integer, nItem As Integer, nPag As Integer
    Dim nTotal As Currency
    
    'Set clsDAdq = New DLogAdquisi
    Set clsDReq = New DLogRequeri
    Set rs = New ADODB.Recordset
    poAdqNro.MoveFirst
    Do While Not poAdqNro.EOF
        nTotal = 0
        nPag = nPag + 1
        sAdqNro = poAdqNro!cAdqNro
        Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroConsul, sAdqNro)
        If rs.RecordCount > 0 Then
            'Inicializacion de Líneas, Items y Páginas
            nLin = 1: nItem = 1
            If poAdqNro.Bookmark > 1 Then sImp = sImp & pCorPag
            sImp = sImp & Cabecera(TpoImpReqAdqui, psAgencia, pdactual, nPag, , sAdqNro)
            'sImp = sImp & ImpreFormat("ADQUISICION N° : " & sAdqNro, 50, 10) & pCorLin & pCorLin
            With rs
                Do While Not .EOF
                    sImp = sImp & ImpreFormat(nItem, 7, 0) & ImpreFormat(!cBsCod, 16) & ImpreFormat(!cBSDescripcion, 30) & ImpreFormat(!cConsUnidad, 6) & _
                        ImpreFormat(!nLogReqDetMesCantidad, 9) & ImpreFormat(!nLogReqDetPrecio, 9) & ImpreFormat(!Subtotal, 10, , True) & pCorLin
                    
                    nTotal = nTotal + !Subtotal
                    nItem = nItem + 1
                    nLin = nLin + 1
                    If nLin > 50 Then
                        sImp = sImp & pCorPag
                        nPag = nPag + 1
                        nLin = 1
                        sImp = sImp & Cabecera(TpoImpReqAdqui, psAgencia, pdactual, nPag, , sAdqNro)
                    End If
                    .MoveNext
                Loop
                sImp = sImp & pCorLin
                sImp = sImp & ImpreFormat("T O T A L", 80, 9) & ImpreFormat(nTotal, 10, , True) & pCorLin
            End With
        End If
        poAdqNro.MoveNext
    Loop
    Set rs = Nothing
    Set clsDReq = Nothing
    
    ImpReqAdqui = sImp
End Function

'Imp. Proveedores
Public Function ImpProveedor(ByVal psAgencia As String, ByVal pdactual As Date, _
Optional ByVal psPersCod As String) As String
    Dim clsDProv As DLogProveedor
    Dim rs As ADODB.Recordset
    
    Dim sImp As String
    Dim nLin As Integer, nItem As Integer, nPag As Integer
    
    Set clsDProv = New DLogProveedor
    Set rs = New ADODB.Recordset
    
    nPag = nPag + 1
    Set rs = clsDProv.CargaProveedor()
    If rs.RecordCount > 0 Then
        'Inicializacion de Líneas, Items y Páginas
        nLin = 1: nItem = 1
        sImp = sImp & Cabecera(TpoImpProveedor, psAgencia, pdactual, nPag)
        With rs
            Do While Not .EOF
                  'INICIO ORCR *****************
                sImp = sImp & ImpreFormat(nItem, 7, 0) & ImpreFormat(!Nombre, 30) & ImpreFormat(!Dirección, 30) & ImpreFormat("" & !RUC, 15) & _
                    ImpreFormat("" & !RUS, 15) & ImpreFormat("" & !DOI, 15) & ImpreFormat(!Estado, 15) & pCorLin
                    'ImpreFormat("" & !RUS, 15) & ImpreFormat("" & !DOI, 15) & ImpreFormat(!Estado, 15) & pCorLin
                  'FIN ORCR *****************
                nItem = nItem + 1
                nLin = nLin + 1
                If nLin > 50 Then
                    sImp = sImp & pCorPag
                    nPag = nPag + 1
                    nLin = 1
                    sImp = sImp & Cabecera(TpoImpProveedor, psAgencia, pdactual, nPag)
                End If
                .MoveNext
            Loop
        End With
    End If
    
    Set rs = Nothing
    Set clsDProv = Nothing
    
    ImpProveedor = sImp
End Function

'Imp. BS
Public Function ImpBS(ByVal psAgencia As String, ByVal pdactual As Date, _
Optional ByVal psPersCod As String) As String
    Dim clsDBS As DLogBieSer
    Dim rs As ADODB.Recordset
    
    Dim sImp As String
    Dim nLin As Integer, nItem As Integer, nPag As Integer
    
    Set clsDBS = New DLogBieSer
    Set rs = New ADODB.Recordset
    
    nPag = nPag + 1
    Set rs = clsDBS.CargaBS(BsTodos)
    If rs.RecordCount > 0 Then
        'Inicializacion de Líneas, Items y Páginas
        nLin = 1: nItem = 1
        sImp = sImp & Cabecera(TpoImpBS, psAgencia, pdactual, nPag)
        With rs
            Do While Not .EOF 'cBSCod , bs.cBSDescripcion, c.cConsDescripcion
                sImp = sImp & ImpreFormat(nItem, 7, 0) & ImpreFormat(!cBsCod, 20) & ImpreFormat(!cBSDescripcion, 60) & ImpreFormat(!cConsDescripcion, 15) & _
                     pCorLin
                
                nItem = nItem + 1
                nLin = nLin + 1
                If nLin > 50 Then
                    sImp = sImp & pCorPag
                    nPag = nPag + 1
                    nLin = 1
                    sImp = sImp & Cabecera(TpoImpBS, psAgencia, pdactual, nPag)
                End If
                .MoveNext
            Loop
        End With
    End If
    
    Set rs = Nothing
    Set clsDBS = Nothing
    
    ImpBS = sImp
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    'Corte de Linea
    pCorLin = oImpresora.gPrnSaltoLinea
    'Corte de Página
    pCorPag = oImpresora.gPrnSaltoPagina
End Sub

Public Function ImpLogPrecios(ByVal psAgencia As String, ByVal pnPeriodo As Integer, pdactual As Date) As String
    Dim clsReq As DLogRequeri
    Dim rs As ADODB.Recordset
    Dim sImp As String
    Dim nLin As Integer, nItem As Integer, nPag As Integer
    Set clsReq = New DLogRequeri
    Set rs = New ADODB.Recordset
    nPag = nPag + 1
    
    
    nLin = 8
    'cBSCod , cBSDescripcion, nLogReqCod, nPrecReferencial
    Set rs = clsReq.CargaListaReqPrecios(pnPeriodo)
    If rs.RecordCount > 0 Then
        'Inicializacion de Líneas, Items y Páginas
        nItem = 1
        sImp = sImp & Cabecera(TpoImpLogPrecios, psAgencia, pdactual, nPag)
        With rs
            'cBSCod,cBSDescripcion,cConsDescripcion,nLogReqCod,nPrecReferencial
            Do While Not .EOF 'cBSCod , bs.cBSDescripcion, c.cConsDescripcion
                sImp = sImp & ImpreFormat(rs!cBsCod, 13) & ImpreFormat(!cBSDescripcion, 40) & ImpreFormat(rs!cConsDescripcion, 8) & ImpreFormat(rs!nLogReqCod, 14) & ImpreFormat(rs!nPrecReferencial, 6) & _
                       pCorLin
                        nItem = nItem + 1
                        nLin = nLin + 1
                        If nLin > 50 Then
                            sImp = sImp & pCorPag
                            nPag = nPag + 1
                            nLin = 1
                            sImp = sImp & Cabecera(TpoImpLogPrecios, psAgencia, pdactual, nPag)
                        End If
                        .MoveNext
            Loop
        End With
    End If
    Set rs = Nothing
    Set clsReq = Nothing
    ImpLogPrecios = sImp
End Function
