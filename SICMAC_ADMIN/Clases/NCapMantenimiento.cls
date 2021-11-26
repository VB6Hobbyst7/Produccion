VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCapMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function GetCTSPeriodo() As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetCTSPeriodo = clsMant.GetCTSPeriodo()
Set clsMant = Nothing
End Function

Public Sub EliminaBeneficiarios(ByVal sPersona As String)
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
clsMant.EliminaBeneficiarios sPersona
Set clsMant = Nothing
End Sub

Public Function ActualizaBeneficiarios(ByVal sPersona As String, ByVal rsBenef As Recordset, ByVal sMovNro As String) As Boolean
Dim clsMant As DCapMantenimiento
Dim bTrans As Boolean
Dim nParentesco As PersRelacion
Set clsMant = New DCapMantenimiento
On Error GoTo ErrGraba
clsMant.dbCmact.BeginTrans
bTrans = True
clsMant.EliminaBeneficiarios sPersona
Do While Not rsBenef.EOF
    nParentesco = CLng(Trim(Right(rsBenef("Parentesco"), 2)))
    clsMant.AgregaBeneficiario sPersona, rsBenef("Codigo"), nParentesco, CDbl(rsBenef("%")), sMovNro
    rsBenef.MoveNext
Loop
clsMant.dbCmact.CommitTrans
bTrans = False
Set clsMant = Nothing
ActualizaBeneficiarios = True
Exit Function
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollBackTrans
    Set clsMant = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
    ActualizaBeneficiarios = False
End Function

Public Function ActualizaTarjetaEstado(ByVal sTarjeta As String, ByVal sMovNro As String, _
        ByVal nEstado As CaptacTarjetaEstado, ByVal sGlosa As String, Optional sClave As String = "") As Boolean

Dim clsMant As DCapMantenimiento
Dim bTrans As Boolean

Set clsMant = New DCapMantenimiento
On Error GoTo ErrGraba
clsMant.dbCmact.BeginTrans
bTrans = True
clsMant.ActualizaTarjetaEstado sTarjeta, nEstado
clsMant.AgregaTarjetaEstado sTarjeta, sMovNro, nEstado, sGlosa
clsMant.dbCmact.CommitTrans
bTrans = False
Set clsMant = Nothing
ActualizaTarjetaEstado = True
Exit Function
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollBackTrans
    Set clsMant = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
    ActualizaTarjetaEstado = False
End Function

Public Function GetOrdenPagoEmitidas(ByVal sCuenta As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetOrdenPagoEmitidas = clsMant.GetOrdenPagoEmitidas(sCuenta)
Set clsMant = Nothing
End Function

Public Function GetDatosPersona(ByVal sPersona As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetDatosPersona = clsMant.GetDatosPersona(sPersona)
Set clsMant = Nothing
End Function

Public Function GetPersonaTarj(ByVal sPersona As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetPersonaTarj = clsMant.GetPersonaTarj(sPersona)
Set clsMant = Nothing
End Function

Public Function GetProductoPersona(ByVal sCuenta As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetProductoPersona = clsMant.GetProductoPersona(sCuenta)
Set clsMant = Nothing
End Function

Public Function GetDatosCuenta(ByVal sCuenta As String) As Recordset
Dim rsCta As Recordset
Dim nProd As Producto
Dim clsMant As DCapMantenimiento
nProd = CInt(Mid(sCuenta, 6, 3))
Set clsMant = New DCapMantenimiento
Select Case nProd
    Case gCapAhorros
        Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
    Case gCapPlazoFijo
        Set rsCta = clsMant.GetDatosCuentaPF(sCuenta)
    Case gCapCTS
        Set rsCta = clsMant.GetDatosCuentaCTS(sCuenta)
End Select
Set GetDatosCuenta = rsCta
Set rsCta = Nothing
Set clsMant = Nothing
End Function

Public Function ActualizaDatosCuenta(ByVal sCuenta As String, ByVal rsRel As Recordset, _
        ByVal nFirmas As Integer, ByVal nTipoCuenta As ProductoCuentaTipo, _
        Optional sInstitucion As String = "", Optional bOrdPag = False) As Boolean

Dim rsCta As Recordset, rsTarj As Recordset
Dim nProd As Producto
Dim clsMant As DCapMantenimiento
Dim bTrans As Boolean
Dim nRelacion As CaptacRelacPersona
Dim sPersona As String
nProd = CInt(Mid(sCuenta, 6, 3))
On Error GoTo ErrGraba
Set clsMant = New DCapMantenimiento
clsMant.dbCmact.BeginTrans
bTrans = True
clsMant.ActualizaDatosCuenta sCuenta, bOrdPag, nFirmas, nTipoCuenta, sInstitucion
clsMant.EliminaProductoPersona sCuenta
Do While Not rsRel.EOF
    sPersona = rsRel("Codigo")
    nRelacion = CLng(Trim(Right(rsRel("Relacion"), 4)))
    clsMant.AgregaProductoPersona sCuenta, sPersona, nRelacion
    rsRel.MoveNext
Loop
clsMant.dbCmact.CommitTrans
Set clsMant = Nothing
ActualizaDatosCuenta = True
Exit Function
ErrGraba:
    If bTrans Then
        clsMant.dbCmact.RollBackTrans
    End If
    Set clsMant = Nothing
    ActualizaDatosCuenta = False
    MsgBox Err.Description, vbExclamation, "Error"
End Function

Public Function GetCuentasPersona(ByVal sPersona As String, Optional nProd As Producto, _
        Optional bActivas As Boolean = False, Optional bBloqueadas As Boolean = False) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetCuentasPersona = clsMant.GetCuentasPersona(sPersona, nProd, bActivas, bBloqueadas)
Set clsMant = Nothing
End Function

Public Function GetCapBloqueos(ByVal sCuenta As String, ByVal nTipoBloqueo As CaptacTipoBloqueo, ByVal nConstante As ConstanteCabecera) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetCapBloqueos = clsMant.GetCapBloqueos(sCuenta, nTipoBloqueo, nConstante)
Set clsMant = Nothing
End Function

Public Sub EliminaProductoPersona(ByVal sCuenta As String)
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
clsMant.EliminaProductoPersona sCuenta
Set clsMant = Nothing
End Sub

Public Function AgregaTarjeta(ByVal sTarjeta As String, ByVal sClave As String, ByVal sMovNro As String, _
        ByVal sPersona As String, ByVal sCuenta As String, ByVal dCaduca As Date) As Boolean
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento

If Not clsMant.ExisteTarjeta(sTarjeta) Then
    clsMant.dbCmact.BeginTrans
    clsMant.AgregaTarjeta sTarjeta, sClave, dCaduca, sMovNro
    clsMant.AgregaCuentaTarj sTarjeta, sCuenta, sPersona
    clsMant.dbCmact.CommitTrans
    AgregaTarjeta = True
Else
    If Not clsMant.ExisteCuentaTarj(sTarjeta, sPersona, sCuenta) Then
        If MsgBox("Tarjeta ya existe. Desea registrar ls misma tarjeta para otra persona??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            clsMant.dbCmact.BeginTrans
            clsMant.AgregaCuentaTarj sTarjeta, sCuenta, sPersona
            clsMant.dbCmact.CommitTrans
            AgregaTarjeta = True
        Else
            AgregaTarjeta = False
        End If
    Else
        MsgBox "Tarjeta ya registrada al cliente. Ingrese una tarjeta nueva", vbInformation, "Aviso"
        AgregaTarjeta = False
    End If
End If
Set clsMant = Nothing
End Function

Public Sub AgregaProductoPersona(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
clsMant.AgregaProductoPersona sCuenta, sPersona, nRelacion
Set clsMant = Nothing
End Sub

Public Sub ModificaProductoPersona(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
'clsMant.ModificaProductoPersona sCuenta, sPersona, nRelacion
Set clsMant = Nothing
End Sub


Public Function GeneraRegistroFirmas(ByVal sCuenta As String, sTipCuenta As String, _
    ByVal dFecApe As Date, ByVal bOrdPag As Boolean, ByVal rsProdPers As Recordset, _
    ByVal sAgencia As String, ByVal dFecSis As Date) As String
    
Dim sTipPer As String
Dim sCodCli As String
Dim sNomCli As String
Dim sDirCli As String
Dim sDocCli As String
Dim sTelCli As String
Dim sNomSec As String
Dim sRazSoc As String
Dim sDirEmp As String
Dim sTelEmp As String
Dim sDocEmp As String
Dim sNumFte As String
Dim sNomAge As String
Dim sDesPro As String * 25, sDesTip As String * 25
Dim sRelCta As String
Dim sClte01 As String
Dim sClte02 As String
Dim sClte03 As String
Dim sClte04 As String
Dim nIngres As Currency
Dim dFchNac As Date
Dim nLinCab As Integer
Dim nCntCli As Integer
Dim I As Integer, nCarLin As Long
Dim sMoneda As String
Dim sCad As String
Dim sNumPag As String
sCad = ""
    
Dim nTamMayor As Long
Dim nTamMenor As Integer
Dim nItem As Integer
Dim nPagina As Integer

Dim sL1 As String * 74
Dim sL2 As String * 74
Dim lsL3 As String * 74
Dim lsL4 As String * 74
Dim lsL5 As String * 74
Dim sTitRp1 As String, sTitRp2 As String
Dim clsMant As DCapMantenimiento
Dim rsPers As Recordset
nTamMayor = 74
nTamMenor = 40
nItem = 12
nPagina = 1
sNomAge = sAgencia
nCarLin = nTamMayor + nTamMenor + 2

Select Case CLng(Mid$(sCuenta, 6, 3))
    Case gCapAhorros
        sDesPro = "AHORRO" & IIf(bOrdPag, " CON ORDEN DE PAGO", "")
    Case gCapPlazoFijo
        sDesPro = "PLAZO FIJO"
    Case gCapCTS
        sDesPro = "C T S"
End Select

sDesTip = sTipCuenta

sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", "NUEVOS SOLES", "DOLARES")
sTitRp1 = "R E G I S T R O   D E   F I R M A S"
sTitRp2 = "CUENTA N° " & Mid(sCuenta, 1, 3) & "-" & Mid(sCuenta, 4, 2) & "-" & Mid(sCuenta, 6, 3) & "-" & Mid(sCuenta, 9, 10)
'sCad = sCad & CabeRepo("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2,  sMoneda, Trim(Str(nPagina)), sAgencia, dFecSis) & oImpresora.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImpresora.gPrnSaltoLinea
sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & oImpresora.gPrnSaltoLinea
sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & oImpresora.gPrnSaltoLinea
sCad = sCad & "FECHA DE APERTURA  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & oImpresora.gPrnSaltoLinea
sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & oImpresora.gPrnSaltoLinea
nLinCab = 5
sCodCli = ""

Set clsMant = New DCapMantenimiento
Set rsPers = New Recordset
rsProdPers.MoveFirst
Do While Not rsProdPers.EOF
    sCodCli = rsProdPers("Codigo")
    Set rsPers = clsMant.GetDatosPersona(sCodCli)
    nCntCli = nCntCli + 1
    sNomCli = ImpreCarEsp(PstaNombre(rsPers("Nombre"), False))
    sRelCta = UCase(Trim(Left(rsProdPers("Relacion"), 30)))
    sDirCli = rsPers("Direccion")
    sDocCli = rsPers("ID N°")
    sTelCli = rsPers("Fono")
    sNomSec = rsPers("Zona")
        
    sL1 = "CODIGO   " & FillText(Trim(sCodCli), 12, " ") & "DNI/LE/CE   " & Trim(sDocCli)
    sL2 = Left(ImpreCarEsp(PstaNombre(sNomCli, False)), nTamMenor)
    lsL3 = Trim(sDirCli) & Space(3) & Trim(sNomSec)
    lsL4 = "TELEFONO   " & sTelCli
    lsL5 = "RC   " & sRelCta
        
    If nItem > 48 Then
        nPagina = nPagina + 1
        sCad = sCad & oImpresora.gPrnSaltoPagina
        'sCad = sCad & CabeRepo("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(Str(sNumPag))) & oImpresora.gPrnSaltoLinea
        sCad = sCad & String(nCarLin, "-") & oImpresora.gPrnSaltoLinea
        sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & oImpresora.gPrnSaltoLinea
        sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & oImpresora.gPrnSaltoLinea
        sCad = sCad & "FECHA DE APERTURA  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & oImpresora.gPrnSaltoLinea
        sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & oImpresora.gPrnSaltoLinea
        nItem = 12
    End If
    
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦") & oImpresora.gPrnSaltoLinea
    sCad = sCad & ImpreCarEsp(sL2 & "¦" & Space(nTamMenor) & "¦") & oImpresora.gPrnSaltoLinea
    sCad = sCad & ImpreCarEsp(lsL3 & "¦" & Space(nTamMenor) & "¦") & oImpresora.gPrnSaltoLinea
    sCad = sCad & ImpreCarEsp(lsL4 & "¦" & Space(nTamMenor) & "¦") & oImpresora.gPrnSaltoLinea
    sCad = sCad & ImpreCarEsp(lsL5 & "¦" & Space(nTamMenor) & "¦") & oImpresora.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & oImpresora.gPrnSaltoLinea
         
    nItem = nItem + 6
    rsProdPers.MoveNext
    rsPers.Close
    Set rsPers = Nothing
Loop

If nItem > 48 Then
    nPagina = nPagina + 1
    sCad = sCad & oImpresora.gPrnSaltoPagina
    'sCad = sCad & CabeRepo("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(Str(sNumPag))) & oImpresora.gPrnSaltoLinea
    sCad = sCad & String(nCarLin, "-") & oImpresora.gPrnSaltoLinea
    sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & oImpresora.gPrnSaltoLinea
    sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & oImpresora.gPrnSaltoLinea
    sCad = sCad & "FECHA DE APERTURA  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & oImpresora.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & oImpresora.gPrnSaltoLinea
    nItem = 12
End If

sL1 = CentrarCadena("INFORME DE ASESORIA LEGAL", nTamMayor)
sL2 = CentrarCadena("FORMAS DE INTERVENCION", nTamMayor)
lsL3 = CentrarCadena("[ ]ABRIR   [ ]CERRAR   [ ]RETIRAR", nTamMayor)
lsL4 = CentrarCadena("______________________________________________________", nTamMayor)

sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦") & oImpresora.gPrnSaltoLinea
sCad = sCad & ImpreCarEsp(sL2 & "¦" & Space(nTamMenor) & "¦") & oImpresora.gPrnSaltoLinea
sCad = sCad & ImpreCarEsp(lsL3 & "¦" & CentrarCadena("_____________________________", nTamMenor - 1) & "¦") & oImpresora.gPrnSaltoLinea
sCad = sCad & ImpreCarEsp(lsL4 & "¦" & CentrarCadena("AUXILIAR DE AHORROS", nTamMenor - 1) & "¦") & oImpresora.gPrnSaltoLinea
sCad = sCad & ImpreCarEsp(lsL4 & "¦" & Space(nTamMenor) & "¦") & oImpresora.gPrnSaltoLinea
sCad = sCad & ImpreCarEsp(lsL4 & "¦" & Space(nTamMenor) & "¦") & oImpresora.gPrnSaltoLinea
sCad = sCad & ImpreCarEsp(lsL4 & "¦" & CentrarCadena("_____________________________", nTamMenor - 1) & "¦") & oImpresora.gPrnSaltoLinea
sCad = sCad & ImpreCarEsp(lsL4 & "¦" & CentrarCadena("ADMINISTRACION DE AHORROS", nTamMenor - 1) & "¦") & oImpresora.gPrnSaltoLinea
sCad = sCad & ImpreCarEsp(lsL4 & "¦" & Space(nTamMenor) & "¦") & oImpresora.gPrnSaltoLinea
sCad = sCad & ImpreCarEsp(lsL4 & "¦" & Space(nTamMenor) & "¦") & oImpresora.gPrnSaltoLinea
sCad = sCad & ImpreCarEsp(lsL4 & "¦" & CentrarCadena("_____________________________", nTamMenor - 1) & "¦") & oImpresora.gPrnSaltoLinea
sCad = sCad & ImpreCarEsp(lsL4 & "¦" & CentrarCadena("ASESOR LEGAL", nTamMenor - 1) & "¦") & oImpresora.gPrnSaltoLinea
sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+"

GeneraRegistroFirmas = sCad
End Function

Public Sub ActualizaBloqueos(ByVal sCuenta As String, ByVal rsRet As Recordset, _
        ByVal rsTot As Recordset, ByVal sMovNro As String, ByVal nEstado As CaptacEstado)

Dim nEstadoCta As CaptacEstado
Dim clsMant As DCapMantenimiento
Dim sComentario As String
Dim nMotRet As CaptacMotBloqueoRet
Dim nMotTot As CaptacMotBloqueoTot
Dim bIniTran As Boolean, bEstado As Boolean

On Error GoTo ErrActBloqueo
Set clsMant = New DCapMantenimiento
clsMant.dbCmact.BeginTrans
bIniTran = True
nEstadoCta = gCapEstActiva
rsRet.MoveFirst
Do While Not rsRet.EOF
    nMotRet = rsRet("nConsValor")
    bEstado = IIf(rsRet("Est") = "0", False, True)
    sComentario = rsRet("Comentario")
    If clsMant.ExisteBloqueoRet(sCuenta, nMotRet) Then
        If Not bEstado Then 'Desbloquea los que ya fueron desmarcados
            clsMant.ActualizaBloqueoRet sCuenta, sComentario, sMovNro, nMotRet
        Else 'Actualiza el comentario de los bloqueos
            clsMant.ActualizaComentBlqRet sCuenta, sComentario, nMotRet
            nEstadoCta = gCapEstBloqRetiro
        End If
    Else
        If bEstado Then 'Agrega los nuevos bloqueos
            clsMant.NuevoBloqueoRetiro sCuenta, nMotRet, sComentario, sMovNro
        End If
    End If
    rsRet.MoveNext
Loop

rsTot.MoveFirst
Do While Not rsTot.EOF
    nMotTot = rsTot("nConsValor")
    bEstado = IIf(rsTot("Est") = "0", False, True)
    sComentario = rsTot("Comentario")
    If clsMant.ExisteBloqueoTot(sCuenta, nMotTot) Then
        If Not bEstado Then 'Agrega los nuevos bloqueos
            clsMant.ActualizaBloqueoTot sCuenta, sComentario, sMovNro, nMotTot
        Else
            clsMant.ActualizaComentBlqTot sCuenta, sComentario, nMotTot
            nEstadoCta = gCapEstBloqTotal
        End If
    Else
        If bEstado Then
            clsMant.NuevoBloqueoTotal sCuenta, nMotTot, sComentario, sMovNro
        End If
    End If
    rsTot.MoveNext
Loop
clsMant.ActualizaEstadoCuenta sCuenta, nEstadoCta
clsMant.dbCmact.CommitTrans
Set clsMant = Nothing
Exit Sub
ErrActBloqueo:
    If bIniTran Then
        clsMant.dbCmact.RollBackTrans
    End If
    Set clsMant = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
End Sub


Public Function GetCapTasaInteres(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
    ByVal nTipoTasa As CaptacTipoTasa, ByVal nValor As Double) As Double

Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
GetCapTasaInteres = clsMant.GetCapTasaInteres(nProducto, nMoneda, nTipoTasa, nValor)
Set clsMant = Nothing
End Function

Public Function GetTasaEfectiva(ByVal nTasaNomAn As Double, ByVal nPeriodo As Integer) As Double
GetTasaEfectiva = (((1 + nTasaNomAn / 36000) ^ nPeriodo) - 1) * 100
End Function

Public Function GetInteresPF(ByVal nTasa As Double, ByVal nCapital As Double, _
                nPlazo As Long) As Double
GetInteresPF = ((1 + nTasa / 36000) ^ nPlazo - 1) * nCapital
End Function

Public Function GetPFPlanRetInt(ByVal dFecApe As Date, ByVal nInteres As Double, nPlazo As Long, nMoneda As Moneda, _
            nIntFinPlazo As Double, nCapital As Double, ByVal nTasa As Double, Optional nIntPriMes As Double = 0, _
            Optional dFecVal As Date) As String

Dim sCad As String
Dim nCarLin As Integer, nDias As Long
Dim sTitRp1 As String, sTitRp2 As String
Dim sMoneda As String, sNumPag As String
Dim nLinPag As Integer, nCntPag As Integer
Dim dFecVenc As Date, dAuxiliar As Date
Dim sFecVenc As String, sFchtra As String
Dim sCodCta As String, sNumDoc As String
Dim sNroCuo As String, nAcum As Double, nUltInteres As Double
Dim bPrimero As Boolean
sCad = ""
sNumPag = ""
nLinPag = 0
nCarLin = 70
sMoneda = IIf(nMoneda = gMonedaNacional, "SOLES", "DOLARES")
sTitRp1 = "CRONOGRAMA  DE  RETIRO  DE INTERESES"
sTitRp2 = ""
sCodCta = ""
nCntPag = 1
nAcum = 0
dAuxiliar = dFecApe
dFecVenc = DateAdd("d", nPlazo, dFecApe)
sFecVenc = Format$(dFecVenc, "dd mmmm yyyy")
sNumPag = FillNum(Trim(Str(nCntPag)), 4, " ")
'sCad = sCad & CabeRepo("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, sNumPag) & oImpresora.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImpresora.gPrnSaltoLinea
sCad = sCad & oImpresora.gPrnSaltoLinea
sCad = sCad & String(3, " ") & "PLAZO             : " & nPlazo & Space(3) & " DIAS" & oImpresora.gPrnSaltoLinea
sCad = sCad & String(3, " ") & "CAPITAL           : " & Format$(nCapital, "#,##0.00") & Space(3) & sMoneda & oImpresora.gPrnSaltoLinea
sCad = sCad & String(3, " ") & "TASA              : " & nTasa & " %" & oImpresora.gPrnSaltoLinea
If nIntPriMes > 0 Then
    sCad = sCad & String(3, " ") & "FECHA APE. CHQ.   : " & Format$(dFecApe, "dd mmmm yyyy") & oImpresora.gPrnSaltoLinea
    sCad = sCad & String(3, " ") & "FECHA VAL. CHQ.   : " & Format$(dFecVal, "dd mmmm yyyy") & oImpresora.gPrnSaltoLinea
Else
    sCad = sCad & String(3, " ") & "FECHA APERTURA    : " & Format$(dFecApe, "dd mmmm yyyy") & oImpresora.gPrnSaltoLinea
End If
sCad = sCad & String(3, " ") & "FECHA VENCIMIENTO : " & sFecVenc & oImpresora.gPrnSaltoLinea
sCad = sCad & String(3, " ") & "MONTO FINAL PLAZO : " & Format$(nIntFinPlazo, "#,##0.00") & " (Sin retirar intereses mensualmente)" & oImpresora.gPrnSaltoLinea
sCad = sCad & oImpresora.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImpresora.gPrnSaltoLinea
sCad = sCad & "  FECHA OPERACION            IMPORTE   " & oImpresora.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImpresora.gPrnSaltoLinea
sCad = sCad & oImpresora.gPrnSaltoLinea
nLinPag = 11
dAuxiliar = DateAdd("d", 30, dAuxiliar)
bPrimero = True
While dFecVenc > dAuxiliar
    sFchtra = Format(dAuxiliar, gsFormatoFechaView)
    If bPrimero Then
        If nIntPriMes > 0 Then
            nAcum = nAcum + nIntPriMes
            'sCad = sCad & "   " & sFchtra & "           " & JDNum(Trim(Str(nIntPriMes)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
        Else
            nAcum = nAcum + nInteres
            'sCad = sCad & "   " & sFchtra & "           " & JDNum(Trim(Str(nInteres)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
        End If
        bPrimero = False
    Else
        'sCad = sCad & "   " & sFchtra & "           " & JDNum(Trim(Str(nInteres)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
        nAcum = nAcum + nInteres
    End If
    
    nLinPag = nLinPag + 1
    
    dAuxiliar = DateAdd("d", 30, dAuxiliar)
Wend
dAuxiliar = DateAdd("d", -30, dAuxiliar)
nDias = DateDiff("d", dAuxiliar, dFecVenc)
If nDias < 30 Then
    nUltInteres = GetInteresPF(nTasa, nCapital, nDias)
    nAcum = nAcum + nUltInteres
    'sCad = sCad & "   " & Format(dFecVenc, gsformatofechaview) & "           " & JDNum(Trim(Str(nUltInteres)), 12, True, 9, 2) & "    RENOVACION" & oImpresora.gPrnSaltoLinea
Else
    nAcum = nAcum + nInteres
    'sCad = sCad & "   " & Format(dFecVenc, gsformatofechaview) & "           " & JDNum(Trim(Str(nInteres)), 12, True, 9, 2) & "    RENOVACION" & oImpresora.gPrnSaltoLinea
End If
sCad = sCad & String(nCarLin, "-") & oImpresora.gPrnSaltoLinea
'sCad = sCad & String(3, " ") & "ACUMULADO " & FillText(sMoneda, 8, " ") & "   " & JDNum(Trim(Str(nAcum)), 12, True, 9, 2) & oImpresora.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "=") & oImpresora.gPrnSaltoLinea
GetPFPlanRetInt = sCad
End Function

Public Function GetBeneficiarios(ByVal sPersona As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetBeneficiarios = clsMant.GetBeneficiarios(sPersona)
Set clsMant = Nothing
End Function

Public Function GetTarjetaCuentas(ByVal sTarjeta As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetTarjetaCuentas = clsMant.GetTarjetaCuentas(sTarjeta)
Set clsMant = Nothing
End Function

Public Function GetTarjetaEstadoHist(ByVal sTarjeta As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetTarjetaEstadoHist = clsMant.GetTarjetaEstadoHist(sTarjeta)
Set clsMant = Nothing
End Function

Public Function ActualizaCuentaTarj(ByVal sTarjeta As String, ByVal sPersona As String, ByVal rsTarj As Recordset) As Boolean
Dim clsMant As DCapMantenimiento
Dim bTrans As Boolean
Dim sCuenta As String
Dim nCuenta As Long
Set clsMant = New DCapMantenimiento
On Error GoTo ErrGrabar
clsMant.dbCmact.BeginTrans
bTrans = True
clsMant.EliminaCuentaTarjPersona sPersona, sTarjeta
Do While Not rsTarj.EOF
    sCuenta = rsTarj("Cuenta")
    clsMant.AgregaCuentaTarjPersona sCuenta, sPersona, sTarjeta
    rsTarj.MoveNext
Loop
clsMant.dbCmact.CommitTrans
Set clsMant = Nothing
ActualizaCuentaTarj = True
Exit Function
ErrGrabar:
    If bTrans Then
        clsMant.dbCmact.RollBackTrans
    End If
    Set clsMant = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
    ActualizaCuentaTarj = False
End Function

Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal nNumOP As Long) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetDatosOrdenPago = clsMant.GetDatosOrdenPago(sCuenta, nNumOP)
Set clsMant = Nothing
End Function

Public Function GetPersonaCuenta(ByVal sCuenta As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetPersonaCuenta = clsMant.GetPersonaCuenta(sCuenta)
Set clsMant = Nothing
End Function

Public Function GeneraConvenioTarjeta(ByVal sCuenta As String, ByVal dFecha As Date) As String
Dim rsPers As Recordset
Dim lsTit As String, lsRep As String
Dim lsDireccion As String, lsDNI As String
Dim sCad As String, sPersona As String, sAux As String
Dim clsMant As DCapMantenimiento

Set clsMant = New DCapMantenimiento
Set rsPers = clsMant.GetPersonaCuenta(sCuenta)
Set clsMant = Nothing

Open App.Path & "\FormatoCarta\ConTar.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, sAux
    sCad = sCad & oImpresora.gPrnSaltoLinea & sAux
Loop
Close #1

If Not (rsPers.EOF And rsPers.BOF) Then
    lsTit = ""
    lsRep = ""
    sPersona = ""
    Do While Not rsPers.EOF
        If sPersona <> rsPers("cPersCod") Then
            If rsPers("nPrdPersRelac") <> gCapRelPersTitular Then
                If lsRep = "" Then
                    lsRep = PstaNombre(Trim(rsPers("Nombre")))
                Else
                    lsRep = lsRep & oImpresora.gPrnSaltoLinea & Space(40) & PstaNombre(Trim(rsPers("Nombre")))
                End If
            Else
                If lsTit = "" Then
                    lsDNI = rsPers("ID") & " " & rsPers("ID N°")
                    lsTit = PstaNombre(rsPers("Nombre"))
                    lsDireccion = rsPers("Direccion")
                Else
                    lsTit = lsTit & oImpresora.gPrnSaltoLinea & Space(52) & PstaNombre(rsPers("Nombre"))
                End If
            End If
            sPersona = rsPers("cPersCod")
        End If
        rsPers.MoveNext
    Loop
    
    sCad = Replace(sCad, "<<NOMBRE>>", lsTit)
    sCad = Replace(sCad, "<<DNI>>", lsDNI)
    sCad = Replace(sCad, "<<DIRECCION>>", lsDireccion)
    sCad = Replace(sCad, "<<REPRESENTANTE>>", lsRep)
    sCad = Replace(sCad, "<<DIA>>", Format$(dFecha, "dddd dd"))
    sCad = Replace(sCad, "<<MES>>", Format$(dFecha, "mmmm"))
    sCad = Replace(sCad, "<<ANIO>>", Format$(dFecha, "yyyy"))
End If
rsPers.Close
Set rsPers = Nothing
GeneraConvenioTarjeta = sCad
End Function


Public Function GetNombreTitulares(ByVal sCuenta As String) As String
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
GetNombreTitulares = clsMant.GetNombreTitulares(sCuenta)
Set clsMant = Nothing
End Function

Public Sub EmiteRangoOrdenPago(ByVal sCuenta As String, ByVal nInicio As Long, _
        ByVal nFin As Long, ByVal sMovNro As String, Optional bDescuento As Boolean = True, _
        Optional nMontoDcto As Double = 0)

Dim sRango As String
Dim bTrans As Boolean
Dim clsMant As DCapMantenimiento
Dim nErrNum As Integer

nErrNum = 0
On Error GoTo ErrGraba
Set clsMant = New DCapMantenimiento

sRango = clsMant.ExisteOrdenPagoEmitidas(sCuenta, nInicio, nFin)
If sRango = "" Then
    If bDescuento Then
        'Agregar codigo para realizar el retiro de la cuenta de ahorros correspondiente
    End If
    clsMant.dbCmact.BeginTrans
    bTrans = True
    clsMant.AgregaOrdenPagoEmitidas sCuenta, nInicio, nFin, sMovNro
    clsMant.dbCmact.CommitTrans
    Set clsMant = Nothing
    Exit Sub
Else
    nErrNum = 1
End If
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollBackTrans
    Set clsMant = Nothing
    If nErrNum = 1 Then
        Err.Raise nErrNum, "EmiteRangoOrdenPago", "Rango ya emitido. Rango Referencia: " & Chr$(13) & sRango
    Else
        Err.Raise Err.Number, "EmiteRangoOrdenPago", Err.Description
    End If
End Sub

Public Sub AnulaRangoOrdenPago(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nNumIni As Long, ByVal nNumFin As Long, ByVal sGlosa As String)

Dim clsMant As DCapMantenimiento
Dim nErrNum As Integer, nAnuladas As Integer
Dim I As Long
Dim rsEstado As Recordset
Dim bTrans As Boolean
Dim nEstado As CaptacOrdPagoEstado
Dim sDesc As String, sMsgErr As String, sEstado As String
Set clsMant = New DCapMantenimiento

On Error GoTo ErrGraba
nErrNum = 0
nAnuladas = 0
sDesc = "Rango Anulado : " & nNumIni & " - " & nNumFin
sGlosa = sGlosa & Chr$(13) & sDesc
clsMant.dbCmact.BeginTrans
bTrans = True
clsMant.AgregaMovNoContable sMovNro, sGlosa, gAhoOPAnulacion
For I = nNumIni To nNumFin
    If clsMant.ExisteOrdenPagoEmitidas(sCuenta, I, I) = "" Then
        sMsgErr = sMsgErr & "Orden Pago NO Emitida : " & I & Chr$(13)
    Else
        Set rsEstado = clsMant.GetEstadoOrdenPagoEmitida(sCuenta, I)
        If rsEstado.EOF And rsEstado.BOF Then
            nEstado = gCapOPEstEmitida
            sEstado = ""
        Else
            nEstado = rsEstado("cEstado")
            sEstado = rsEstado("Estado")
        End If
        If nEstado <> gCapOPEstEmitida And nEstado <> gCapOPEstExtornada Then
            sMsgErr = sMsgErr & "Orden Pago NO Válida : " & I & Space(2) & sEstado & Chr$(13)
        Else
            If nEstado = gCapOPEstEmitida Then clsMant.AgregaOrdenPagoRecibida sCuenta, I, 0
            clsMant.AnulaOrdenPagoEmitida sCuenta, I, sMovNro, 0
            nAnuladas = nAnuladas + 1
        End If
    End If
Next I
If nAnuladas = 0 Then
    clsMant.dbCmact.RollBackTrans
Else
    clsMant.dbCmact.CommitTrans
End If
bTrans = False
Set clsMant = Nothing
If sMsgErr = "" Then
    Exit Sub
Else
    nErrNum = 1
End If
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollBackTrans
    Set clsMant = Nothing
    If nErrNum = 0 Then
        Err.Raise Err.Number, "AnulaRangoOrdenPago", Err.Description
    ElseIf nErrNum = 1 Then
        Err.Raise nErrNum, "AnulaRangoOrdenPago", "Ordenes Anuladas : " & nAnuladas & Chr$(13) & sMsgErr
    End If
End Sub

Public Sub AnulaOrdenPago(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nNumOP As Long, ByVal sGlosa As String, Optional nMonto As Double = 0)

Dim clsMant As DCapMantenimiento
Dim nErrNum As Integer
Dim rsEstado As Recordset
Dim bTrans As Boolean
Set clsMant = New DCapMantenimiento
Dim sEstado As String
Dim nEstado As CaptacOrdPagoEstado
On Error GoTo ErrGraba
nErrNum = 0
If clsMant.ExisteOrdenPagoEmitidas(sCuenta, nNumOP, nNumOP) = "" Then
    nErrNum = 1
    GoTo ErrGraba
End If
Set rsEstado = clsMant.GetEstadoOrdenPagoEmitida(sCuenta, nNumOP)
If rsEstado.EOF And rsEstado.BOF Then
    nEstado = gCapOPEstEmitida
    sEstado = ""
Else
    nEstado = rsEstado("cEstado")
    sEstado = rsEstado("Estado")
End If
If nEstado <> gCapOPEstEmitida And nEstado <> gCapOPEstExtornada Then
    nErrNum = 2
    GoTo ErrGraba
End If
clsMant.dbCmact.BeginTrans
bTrans = True
If nEstado = gCapOPEstEmitida Then clsMant.AgregaOrdenPagoRecibida sCuenta, nNumOP, nMonto
clsMant.AgregaMovNoContable sMovNro, sGlosa, gAhoOPAnulacion
clsMant.AnulaOrdenPagoEmitida sCuenta, nNumOP, sMovNro, nMonto
clsMant.dbCmact.CommitTrans
bTrans = False
Set clsMant = Nothing
Exit Sub
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollBackTrans
    Set clsMant = Nothing
    If nErrNum = 0 Then
        Err.Raise Err.Number, "AnulaOrdenPago", Err.Description
    ElseIf nErrNum = 1 Then
        Err.Raise nErrNum, "AnulaOrdenPago", "Orden Pago NO Emitida"
    ElseIf nErrNum = 2 Then
        Err.Raise nErrNum, "AnulaOrdenPago", "Orden Pago NO Válida. Estado " & sEstado
    End If
End Sub

Public Function GetSaldoFecha(ByVal sCuenta As String, ByVal dFecha As Date) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetSaldoFecha = clsMant.GetSaldoFecha(sCuenta, dFecha)
Set clsMant = Nothing
End Function


Public Function GetMovimientosCuenta(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetMovimientosCuenta = clsMant.GetMovimientosCuenta(sCuenta, dFecIni, dFecFin, nNumMov)
Set clsMant = Nothing
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing


End Sub
