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
    TpoImpAdquisicion = 1
    TpoImpProveedor = 2
    TpoImpBS = 3
End Enum

Private Function Cabecera(ByVal psIMP As TpoImp, ByVal psAge As String, ByVal pdAct As Date, _
ByVal pnPag As Integer, Optional ByVal pbPagCorta As Boolean = True, _
Optional ByVal psAdqNro As String = "") As String
    Dim sCabe As String
    Dim sTit As String, sTitSub As String
    Dim sArea As String
    Dim nNroCol As Integer
    
    sTit = "": sTitSub = ""
    Select Case psIMP
        Case TpoImpAdquisicion
            sTit = "L I S T A D O   D E   A D Q U I S I C I O N E S"
            sTitSub = "Adquisición : " & psAdqNro
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
    sCabe = sCabe & Space(1) & ImpreFormat(sArea, 25, 0) & sTitSub & Space(4) & Format(pdAct & " " & Time, "dd/mm/yyyy hh:mm:ss") & pCorLin
    
    sCabe = sCabe & String(nNroCol, "-") & pCorLin
    Select Case psIMP
        Case TpoImpAdquisicion
            sCabe = sCabe & Space(2) & " ITEM      CODIGO                BIEN/SERVICIO           UNIDAD    CANTIDAD      PRECIO     SUBTOTAL" & pCorLin
        Case TpoImpProveedor
            sCabe = sCabe & Space(2) & " ITEM             NOMBRE                         DIRECCION                 RUC              RUS            ESTADO" & pCorLin
        Case TpoImpBS
            sCabe = sCabe & Space(2) & " ITEM       CODIGO                             NOMBRE                                      UNIDAD   " & pCorLin
    End Select
    sCabe = sCabe & String(nNroCol, "-") & pCorLin
    
    Cabecera = sCabe
End Function


'Imp. Adquisiciones
Public Function ImpAdquisicion(ByVal psAgencia As String, ByVal pdActual As Date, _
ByVal poAdqNro As ADODB.Recordset) As String
    Dim clsDAdq As DLogAdquisi
    Dim Rs As ADODB.Recordset
    Dim sAdqNro As String
    Dim sImp As String
    Dim nLin As Integer, nItem As Integer, nPag As Integer
    Dim nTotal As Currency
    
    Set clsDAdq = New DLogAdquisi
    Set Rs = New ADODB.Recordset
    poAdqNro.MoveFirst
    Do While Not poAdqNro.EOF
        nTotal = 0
        nPag = nPag + 1
        sAdqNro = poAdqNro!cAdqNro
        Set Rs = clsDAdq.CargaAdqDetalle(AdqDetUnRegistro, sAdqNro)
        If Rs.RecordCount > 0 Then
            'Inicializacion de Líneas, Items y Páginas
            nLin = 1: nItem = 1
            If poAdqNro.Bookmark > 1 Then sImp = sImp & pCorPag
            sImp = sImp & Cabecera(TpoImpAdquisicion, psAgencia, pdActual, nPag, , sAdqNro)
            'sImp = sImp & ImpreFormat("ADQUISICION N° : " & sAdqNro, 50, 10) & pCorLin & pCorLin
            With Rs
                Do While Not .EOF
                    sImp = sImp & ImpreFormat(nItem, 7, 0) & ImpreFormat(!cBScod, 16) & ImpreFormat(!cBSDescripcion, 30) & ImpreFormat(!cConsUnidad, 6) & _
                        ImpreFormat(!nLogAdqDetCantidad, 9) & ImpreFormat(!nLogAdqDetPrecio, 9) & ImpreFormat(!SubTotal, 10, , True) & pCorLin
                    
                    nTotal = nTotal + !SubTotal
                    nItem = nItem + 1
                    nLin = nLin + 1
                    If nLin > 50 Then
                        sImp = sImp & pCorPag
                        nPag = nPag + 1
                        nLin = 1
                        sImp = sImp & Cabecera(TpoImpAdquisicion, psAgencia, pdActual, nPag, , sAdqNro)
                    End If
                    .MoveNext
                Loop
                sImp = sImp & pCorLin
                sImp = sImp & ImpreFormat("T O T A L", 80, 9) & ImpreFormat(nTotal, 10, , True) & pCorLin
            End With
        End If
        poAdqNro.MoveNext
    Loop
    Set Rs = Nothing
    Set clsDAdq = Nothing
    
    ImpAdquisicion = sImp
End Function

'Imp. Proveedores
Public Function ImpProveedor(ByVal psAgencia As String, ByVal pdActual As Date, _
Optional ByVal psPersCod As String) As String
    Dim clsDProv As DLogProveedor
    Dim Rs As ADODB.Recordset
    
    Dim sImp As String
    Dim nLin As Integer, nItem As Integer, nPag As Integer
    
    Set clsDProv = New DLogProveedor
    Set Rs = New ADODB.Recordset
    
    nPag = nPag + 1
    Set Rs = clsDProv.CargaProveedor()
    If Rs.RecordCount > 0 Then
        'Inicializacion de Líneas, Items y Páginas
        nLin = 1: nItem = 1
        sImp = sImp & Cabecera(TpoImpProveedor, psAgencia, pdActual, nPag)
        With Rs
            Do While Not .EOF
                sImp = sImp & ImpreFormat(nItem, 7, 0) & ImpreFormat(!Nombre, 30) & ImpreFormat(!Dirección, 30) & ImpreFormat("" & !RUC, 15) & _
                    ImpreFormat("" & !RUS, 15) & ImpreFormat(!Estado, 15) & pCorLin
                
                nItem = nItem + 1
                nLin = nLin + 1
                If nLin > 50 Then
                    sImp = sImp & pCorPag
                    nPag = nPag + 1
                    nLin = 1
                    sImp = sImp & Cabecera(TpoImpProveedor, psAgencia, pdActual, nPag)
                End If
                .MoveNext
            Loop
        End With
    End If
    
    Set Rs = Nothing
    Set clsDProv = Nothing
    
    ImpProveedor = sImp
End Function

'Imp. BS
Public Function ImpBS(ByVal psAgencia As String, ByVal pdActual As Date, _
Optional ByVal psPersCod As String) As String
    Dim clsDBS As DLogBieSer
    Dim Rs As ADODB.Recordset
    
    Dim sImp As String
    Dim nLin As Integer, nItem As Integer, nPag As Integer
    
    Set clsDBS = New DLogBieSer
    Set Rs = New ADODB.Recordset
    
    nPag = nPag + 1
    Set Rs = clsDBS.CargaBS(BsTodos)
    If Rs.RecordCount > 0 Then
        'Inicializacion de Líneas, Items y Páginas
        nLin = 1: nItem = 1
        sImp = sImp & Cabecera(TpoImpBS, psAgencia, pdActual, nPag)
        With Rs
            Do While Not .EOF 'cBSCod , bs.cBSDescripcion, c.cConsDescripcion
                sImp = sImp & ImpreFormat(nItem, 7, 0) & ImpreFormat(!cBScod, 20) & ImpreFormat(!cBSDescripcion, 60) & ImpreFormat(!cConsDescripcion, 15) & _
                     pCorLin
                
                nItem = nItem + 1
                nLin = nLin + 1
                If nLin > 50 Then
                    sImp = sImp & pCorPag
                    nPag = nPag + 1
                    nLin = 1
                    sImp = sImp & Cabecera(TpoImpBS, psAgencia, pdActual, nPag)
                End If
                .MoveNext
            Loop
        End With
    End If
    
    Set Rs = Nothing
    Set clsDBS = Nothing
    
    ImpBS = sImp
End Function

Private Sub Class_Initialize()
    'Corte de Linea
    pCorLin = Chr(10)
    'Corte de Página
    pCorPag = Chr(12)
End Sub
