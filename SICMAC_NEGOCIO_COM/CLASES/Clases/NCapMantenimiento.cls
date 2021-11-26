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
    If bTrans Then clsMant.dbCmact.RollbackTrans
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
clsMant.ActualizaTarjetaEstado sTarjeta, nEstado, sClave
clsMant.AgregaTarjetaEstado sTarjeta, sMovNro, nEstado, sGlosa
clsMant.dbCmact.CommitTrans
bTrans = False
Set clsMant = Nothing
ActualizaTarjetaEstado = True
Exit Function
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
    ActualizaTarjetaEstado = False
End Function

Public Function GetCuentasNoAsociadas(ByVal sPerscod As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetCuentasNoAsociadas = clsMant.GetCuentasNoAsociadas(sPerscod)
Set clsMant = Nothing
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


Public Function GetCuentasPersonaAI(Optional ByVal sPerscod As String) As Recordset
  Dim clsMant As DCapMantenimiento, RSTEMP As Recordset
  Set clsMant = New DCapMantenimiento
  
              Set RSTEMP = clsMant.GetCuentasPersonaAI(sPerscod)
              Set GetCuentasPersonaAI = RSTEMP
  
  Set clsMant = Nothing
End Function

Public Function GetTpoActI() As Recordset
  Dim clsMant As DCapMantenimiento, RSTEMP As Recordset
  Set clsMant = New DCapMantenimiento
  
              Set RSTEMP = clsMant.GetTpoActI()
              Set GetTpoActI = RSTEMP
  
  Set clsMant = Nothing
End Function

Public Function GetTpoDocActI() As Recordset
  Dim clsMant As DCapMantenimiento, RSTEMP As Recordset
  Set clsMant = New DCapMantenimiento
  
              Set RSTEMP = clsMant.GetTpoDocActI()
              Set GetTpoDocActI = RSTEMP
  
  Set clsMant = Nothing
End Function

Public Function GetDPTO() As Recordset
  Dim clsMant As DCapMantenimiento, RSTEMP As Recordset
  Set clsMant = New DCapMantenimiento
  
              Set RSTEMP = clsMant.GetDPTO()
              Set GetDPTO = RSTEMP
  
  Set clsMant = Nothing
End Function

Public Function GetPROV(ByVal cDpto As String) As Recordset
  Dim clsMant As DCapMantenimiento, RSTEMP As Recordset
  Set clsMant = New DCapMantenimiento
  
              Set RSTEMP = clsMant.GetPROV(cDpto)
              Set GetPROV = RSTEMP
  
  Set clsMant = Nothing
End Function



Public Function EsCtaConvenio(ByVal sCuenta As String) As Boolean
EsCtaConvenio = False
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
    
     
    EsCtaConvenio = clsMant.EsCtaConvenio(sCuenta)
    
    Set clsMant = Nothing

End Function


Public Function ActualizaDatosCuenta(ByVal sCuenta As String, ByVal rsRel As Recordset, _
        ByVal nFirmas As Integer, ByVal nTipoCuenta As ProductoCuentaTipo, _
        Optional sInstitucion As String = "", Optional bOrdPag As Boolean = False, _
        Optional sCuentaAboInt As String = "", Optional nFirmasMin As Integer, Optional sMovNro As String = "", _
        Optional rsEliminados As Recordset, Optional sAlias As String = "", _
        Optional ByVal nTipoPrograma As Integer = 0) As Boolean

Dim rsCta As Recordset, rsTarj As Recordset
Dim nProd As Producto
Dim clsMant As DCapMantenimiento
Dim bTrans As Boolean, COBLIGATORIO As String
Dim nRelacion As CaptacRelacPersona
Dim sPersona As String
nProd = CInt(Mid(sCuenta, 6, 3))
On Error GoTo ErrGraba
Set clsMant = New DCapMantenimiento
clsMant.dbCmact.BeginTrans
bTrans = True
clsMant.ActualizaDatosCuenta sCuenta, bOrdPag, nFirmas, nTipoCuenta, sInstitucion, sCuentaAboInt, nFirmasMin, sAlias, nTipoPrograma
clsMant.EliminaProductoPersona sCuenta

Do While Not rsRel.EOF
    sPersona = rsRel("Codigo")
    nRelacion = CLng(Trim(Right(rsRel("Relacion"), 4)))
    COBLIGATORIO = Left(rsRel("Obligatorio"), 1)
    clsMant.AgregaProductoPersona sCuenta, sPersona, nRelacion, COBLIGATORIO
    clsMant.AgregaBitacoraPPer sMovNro, sCuenta, sPersona, nRelacion, COBLIGATORIO
    
    rsRel.MoveNext
Loop

clsMant.dbCmact.CommitTrans
Set clsMant = Nothing
ActualizaDatosCuenta = True
Exit Function
ErrGraba:
    If bTrans Then
        clsMant.dbCmact.RollbackTrans
    End If
    Set clsMant = Nothing
    ActualizaDatosCuenta = False
    MsgBox Err.Description, vbExclamation, "Error"
End Function

Public Function GetCuentasPersona(ByVal sPersona As String, Optional nProd As Producto, _
        Optional bActivas As Boolean = False, Optional bBloqueadas As Boolean = False, _
        Optional nMoneda As Moneda, Optional bOrdPag As Boolean = False, _
        Optional sAgencia As String = "") As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetCuentasPersona = clsMant.GetCuentasPersona(sPersona, nProd, bActivas, bBloqueadas, nMoneda, bOrdPag, sAgencia)
Set clsMant = Nothing
End Function
Public Function GetCuentasAhorro(ByVal bTodas As Boolean, Optional cAgeCod As String = "", Optional psCctacod As String = "") As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetCuentasAhorro = clsMant.GetCuentasAhorro(bTodas, cAgeCod, psCctacod)
Set clsMant = Nothing

End Function
Public Sub ActualizaCtaDescInact(ByVal psCuenta As String, ByVal bdescinactiva As Boolean, ByVal psMovNro As String)
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
       Call clsMant.ActualizaCtaDescInact(psCuenta, bdescinactiva, psMovNro)
       Set clsMant = Nothing
End Sub



Public Function GetCapBloqueos(ByVal sCuenta As String, ByVal nTipoBloqueo As CaptacTipoBloqueo, ByVal nConstante As ConstanteCabecera) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetCapBloqueos = clsMant.GetCapBloqueos(sCuenta, nTipoBloqueo, nConstante)
Set clsMant = Nothing
End Function

Public Sub EliminaProductoPersona(ByVal sCuenta As String, ByVal sPersona As String)
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
    MsgBox "Tarjeta ya registrada al cliente. Ingrese una tarjeta nueva", vbInformation, "Aviso"
    AgregaTarjeta = False
End If
Set clsMant = Nothing
End Function


Public Function AgregaTarjetaLOTE(ByVal sTarjeta As String, ByVal sClave As String, ByVal sMovNro As String, _
         ByVal rCuentas As Recordset, ByVal dCaduca As Date, ByVal sPerscod As String) As Boolean
         
Dim clsMant As DCapMantenimiento, rsPers As Recordset
Set clsMant = New DCapMantenimiento


If Not clsMant.ExisteTarjeta(sTarjeta) Then
    
    If Not clsMant.EsTarjetaEmitida(sTarjeta) Then
        MsgBox "Esta Tarjeta no ha sido emitida por la CMAC ICA", vbOKOnly + vbExclamation, "AVISO IMPORTANTE"
        AgregaTarjetaLOTE = False
        Exit Function
    End If


    clsMant.dbCmact.BeginTrans
    
     clsMant.AgregaTarjeta sTarjeta, sClave, dCaduca, sMovNro
        
     While Not rCuentas.EOF
        'Set rsPers = New Recordset
        'Set rsPers = clsmant.GetPersonasCuentaNA(rCuentas.Fields("Cuenta"), "", CInt(rCuentas.Fields("Personeria")))
         
        ' While Not rsPers.EOF
           ' clsmant.AgregaCuentaTarj sTarjeta, rCuentas.Fields("Cuenta"), rsPers.Fields("Codigo")
           ' rsPers.MoveNext
        ' Wend
         clsMant.AgregaCuentaTarj sTarjeta, rCuentas.Fields("Cuenta"), sPerscod
         
       ' Set rsPers = Nothing
        rCuentas.MoveNext
     Wend
    
    clsMant.dbCmact.CommitTrans
    AgregaTarjetaLOTE = True
Else
    MsgBox "Tarjeta ya registrada al cliente. Ingrese una tarjeta nueva", vbInformation, "Aviso"
    AgregaTarjetaLOTE = False
End If

Set clsMant = Nothing
End Function


Public Sub AgregaProductoPersona(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
clsMant.AgregaProductoPersona sCuenta, sPersona, nRelacion
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
Dim i As Integer, nCarLin As Long
Dim sMoneda As String
Dim sCad As String
Dim sNumPag As String
sCad = ""
    
Dim nTamMayor As Long, nTamMenor As Long, nTamAdd As Long
Dim nitem As Integer
Dim nPagina As Integer

Dim sL1 As String * 74
Dim sL2 As String * 74
Dim sL3 As String * 74
Dim sL4 As String * 74
Dim sL5 As String * 74
Dim sTitRp1 As String, sTitRp2 As String
Dim clsMant As DCapMantenimiento
Dim rsPers As Recordset
nTamMayor = 74
nTamMenor = 40
nTamAdd = 22
nitem = 12
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
sCad = sCad & CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), sAgencia, dFecSis) & Chr$(10)
sCad = sCad & String(nTamMayor + nTamMenor + 2 + nTamAdd, "-") & Chr$(10)
sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & Chr(10)
sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & Chr$(10)
sCad = sCad & "FECHA DE APERTURA  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & Space(35) & "HUELLA" & Chr$(10)
sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & String(nTamAdd, "-") & "+" & Chr$(10)
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
    sDirCli = rsPers("Direccion") & ""
    sDocCli = rsPers("ID N°")
    sTelCli = rsPers("Fono")
    sNomSec = rsPers("Zona")
        
    sL1 = "CODIGO   " & FillText(Trim(sCodCli), 13, " ") & "DNI/LE/CE   " & Trim(sDocCli)
    sL2 = Left(ImpreCarEsp(PstaNombre(sNomCli, False)), nTamMenor)
    sL3 = Trim(sDirCli) & Space(3) & Trim(sNomSec)
    sL4 = "TELEFONO   " & sTelCli
    sL5 = "RC   " & sRelCta
        
    If nitem > 48 Then
        nPagina = nPagina + 1
        sCad = sCad & Chr(12)
        sCad = sCad & CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(sNumPag)), "", dFecSis) & Chr$(10)
        sCad = sCad & String(nTamMayor + nTamMenor + 2 + nTamAdd, "-") & Chr$(10)
        sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & Chr$(10)
        sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & Chr$(10)
        sCad = sCad & "FECHA DE APERTURA  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & Space(35) & "HUELLA" & Chr$(10)
        sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & String(nTamAdd, "-") & "+" & Chr$(10)
        nitem = 12
    End If
    
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL2 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL3 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL4 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL5 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    
    sL1 = ""
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & CentrarCadena("..........................................", nTamMenor - 1) & " ¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-" & "+") & "+" & String(nTamAdd, "-") & "+" & Chr$(10)
         
    nitem = nitem + 12
    rsProdPers.MoveNext
    rsPers.Close
    Set rsPers = Nothing
Loop
Dim arreglocad As String
If nitem > 48 Then
    nPagina = nPagina + 1
    sCad = sCad & Chr(12)
    sCad = sCad & "" 'CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(Str(sNumPag)), "", dFecSis) & Chr$(10)
    arreglocad = CabeRepoCaptac(" ", " ", 1, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(sNumPag)), "", dFecSis)
    sCad = sCad & arreglocad & Chr$(10)
    sCad = sCad & String(nCarLin, "-") & Chr$(10)
    sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & Chr(10)
    sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & Chr$(10)
    sCad = sCad & "FECHA DE APERTURA  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & Chr(10)
    sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & Chr(10)
    nitem = 12
End If

sL1 = CentrarCadena("INFORME DE ASESORIA LEGAL", nTamMayor)
sL2 = CentrarCadena(" FORMAS DE INTERVENCION", nTamMayor)
sL3 = CentrarCadena("[ ]ABRIR   [ ]CERRAR   [ ]RETIRAR", nTamMayor)
sL4 = ""

sCad = sCad & ImpreCarEsp(CentrarCadena(sL1, nTamMayor + nTamMenor + nTamAdd) & " ¦") & Chr$(10)
sCad = sCad & ImpreCarEsp(CentrarCadena(sL2, nTamMayor + nTamMenor + nTamAdd) & "  ¦") & Chr$(10)
sCad = sCad & ImpreCarEsp(CentrarCadena(sL3, nTamMayor + nTamMenor + nTamAdd) & " ¦") & Chr$(10)
sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & "  ¦" & Chr$(10)
sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & "  ¦" & Chr$(10)
sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & "  ¦" & Chr$(10)
sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & "  ¦" & Chr$(10)
sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & "  ¦" & Chr$(10)
sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & "  ¦" & Chr$(10)
sCad = sCad & String(nTamMayor, "-") & "-" & String(nTamMenor + nTamAdd, "-") & "-+" & Chr$(10)

sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & " ¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & " ¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & " ¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & " ¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena("_____________________________               _____________________________             _____________________________", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena("    AUXILIAR DE AHORROS                       ASISTENTE DE OPERACIONES                       ASESORIA LEGAL        ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & " ¦") & Chr$(10)
sCad = sCad & String(nTamMayor, "-") & "-" & String(nTamMenor + nTamAdd, "-") & "-+"
sCad = sCad & Chr(12)
sCad = sCad & RegistroFirmasCANC(sCuenta, bOrdPag, rsProdPers, dFecSis)
GeneraRegistroFirmas = sCad
End Function
Public Function RegistroFirmasCANC(ByVal sCuenta As String, ByVal bOrdPag As Boolean, ByVal rsProdPers As Recordset, Optional pdFecSis As Date = CDate("01/01/1900")) As String
    
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
Dim i As Integer, nCarLin As Long
Dim sMoneda As String
Dim sCad As String
Dim sNumPag As String
sCad = ""
    
Dim nTamMayor As Long, nTamMenor As Long, nTamAdd As Long
Dim nitem As Integer
Dim nPagina As Integer

Dim sL1 As String * 74
Dim sL2 As String * 74
Dim sL3 As String * 74
Dim sL4 As String * 74
Dim sL5 As String * 74
Dim sTitRp1 As String, sTitRp2 As String
Dim clsMant As DCapMantenimiento
Dim rsPers As Recordset
nTamMayor = 74
nTamMenor = 40
nTamAdd = 22
nitem = 12
nPagina = 1
'sNomAge = sAgencia
nCarLin = nTamMayor + nTamMenor + 2

Select Case CLng(Mid$(sCuenta, 6, 3))
    Case gCapAhorros
        sDesPro = "AHORRO" & IIf(bOrdPag, " CON ORDEN DE PAGO", "")
    Case gCapPlazoFijo
        sDesPro = "PLAZO FIJO"
    Case gCapCTS
        sDesPro = "C T S"
End Select

'sDesTip = sTipCuenta

If pdFecSis = "01/01/1900" Then
   pdFecSis = Date
End If

sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", "NUEVOS SOLES", "DOLARES")
sTitRp1 = "R E G I S T R O   D E   F I R M A S  -  C A N C E L A C I O N   D E   C U E N T A  "
sTitRp2 = "CUENTA N° " & Mid(sCuenta, 1, 3) & "-" & Mid(sCuenta, 4, 2) & "-" & Mid(sCuenta, 6, 3) & "-" & Mid(sCuenta, 9, 10)
sCad = sCad & CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), "", pdFecSis) & Chr$(10)
sCad = sCad & String(nTamMayor + nTamMenor + 2 + nTamAdd, "-") & Chr$(10)
sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & "_______________________" & Chr(10)
sCad = sCad & "FECHA DE CANCELACION  : " & "__________" & Space(60) & "FIRMA" & Space(35) & "HUELLA" & Chr$(10)
sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & String(nTamAdd, "-") & "+" & Chr$(10)
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
    sDirCli = rsPers("Direccion") & ""
    sDocCli = rsPers("ID N°")
    sTelCli = rsPers("Fono")
    sNomSec = rsPers("Zona")
        
    sL1 = "CODIGO   " & FillText(Trim(sCodCli), 13, " ") & "DNI/LE/CE   " & Trim(sDocCli)
    sL2 = Left(ImpreCarEsp(PstaNombre(sNomCli, False)), nTamMenor)
    sL3 = Trim(sDirCli) & Space(3) & Trim(sNomSec)
    sL4 = "TELEFONO   " & sTelCli
    sL5 = "RC   " & sRelCta
        
    If nitem > 48 Then
        nPagina = nPagina + 1
        sCad = sCad & Chr(12)
        sCad = sCad & CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), "", pdFecSis) & Chr$(10)
        sCad = sCad & String(nTamMayor + nTamMenor + 2 + nTamAdd, "-") & Chr$(10)
        sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & "_______________________" & Chr(10)
        sCad = sCad & "FECHA DE CANCELACION  : " & "__________" & Space(60) & "FIRMA" & Space(35) & "HUELLA" & Chr$(10)
        sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & String(nTamAdd, "-") & "+" & Chr$(10)
        nitem = 12
    End If
    
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL2 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL3 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL4 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL5 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    
    sL1 = ""
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & CentrarCadena(".........................................", nTamMenor - 1) & " ¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-" & "+") & "+" & String(nTamAdd, "-") & "+" & Chr$(10)
         
    nitem = nitem + 12
    rsProdPers.MoveNext
    rsPers.Close
    Set rsPers = Nothing
Loop
Dim arreglocad As String
'If nItem > 48 Then
'    nPagina = nPagina + 1
'    sCad = sCad & Chr(12)
'    sCad = sCad & "" 'CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(Str(sNumPag)), "", dFecSis) & Chr$(10)
'    arreglocad = CabeRepoCaptac(" ", " ", 1, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(sNumPag)), "", dFecSis)
'    sCad = sCad & arreglocad & Chr$(10)
'    sCad = sCad & String(nCarLin, "-") & Chr$(10)
'    sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & Chr(10)
'    sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & Chr$(10)
'    sCad = sCad & "FECHA DE CANCELACION  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & Chr(10)
'    sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & Chr(10)
'    nItem = 12
'End If
'
'sL1 = CentrarCadena("INFORME DE ASESORIA LEGAL", nTamMayor)
'sL2 = CentrarCadena("FORMAS DE INTERVENCION", nTamMayor)
'sL3 = CentrarCadena("[ ]ABRIR   [ ]CERRAR   [ ]RETIRAR", nTamMayor)
'sL4 = ""
'
'sCad = sCad & ImpreCarEsp(CentrarCadena(sL1, nTamMayor + nTamMenor + nTamAdd) & " ¦") & Chr$(10)
'sCad = sCad & ImpreCarEsp(CentrarCadena(sL2, nTamMayor + nTamMenor + nTamAdd) & " ¦") & Chr$(10)
'sCad = sCad & ImpreCarEsp(CentrarCadena(sL3, nTamMayor + nTamMenor + nTamAdd) & " ¦") & Chr$(10)
'sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & Chr$(10)
'sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & Chr$(10)
'sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & Chr$(10)
'sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & Chr$(10)
'sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & Chr$(10)
'sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & Chr$(10)
'sCad = sCad & String(nTamMayor, "-") & "-" & String(nTamMenor + nTamAdd, "-") & "-+" & Chr$(10)
'
'sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
'sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
'sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
'sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
'sCad = sCad & ImpreCarEsp("¦" & CentrarCadena("_____________________________               _____________________________             _____________________________", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
'sCad = sCad & ImpreCarEsp("¦" & CentrarCadena("    AUXILIAR DE AHORROS                       ASISTENTE DE OPERACIONES                       ASESORIA LEGAL        ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
sCad = sCad & String(nTamMayor, "-") & "-" & String(nTamMenor + nTamAdd, "-") & "-+"
    
    RegistroFirmasCANC = sCad
End Function



Public Function GeneraRegistroFirmasCANCELACION(ByVal sCuenta As String, sTipCuenta As String, _
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
Dim i As Integer, nCarLin As Long
Dim sMoneda As String
Dim sCad As String
Dim sNumPag As String
sCad = ""
    
Dim nTamMayor As Long, nTamMenor As Long, nTamAdd As Long
Dim nitem As Integer
Dim nPagina As Integer

Dim sL1 As String * 74
Dim sL2 As String * 74
Dim sL3 As String * 74
Dim sL4 As String * 74
Dim sL5 As String * 74
Dim sTitRp1 As String, sTitRp2 As String
Dim clsMant As DCapMantenimiento
Dim rsPers As Recordset
nTamMayor = 74
nTamMenor = 40
nTamAdd = 22
nitem = 12
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
sCad = sCad & CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), sAgencia, dFecSis) & Chr$(10)
sCad = sCad & String(nTamMayor + nTamMenor + 2 + nTamAdd, "-") & Chr$(10)
sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & Chr(10)
sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & Chr$(10)
sCad = sCad & "FECHA DE CANCELACION  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & Space(35) & "HUELLA" & Chr$(10)
sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & String(nTamAdd, "-") & "+" & Chr$(10)
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
    sDirCli = rsPers("Direccion") & ""
    sDocCli = rsPers("ID N°")
    sTelCli = rsPers("Fono")
    sNomSec = rsPers("Zona")
        
    sL1 = "CODIGO   " & FillText(Trim(sCodCli), 13, " ") & "DNI/LE/CE   " & Trim(sDocCli)
    sL2 = Left(ImpreCarEsp(PstaNombre(sNomCli, False)), nTamMenor)
    sL3 = Trim(sDirCli) & Space(3) & Trim(sNomSec)
    sL4 = "TELEFONO   " & sTelCli
    sL5 = "RC   " & sRelCta
        
    If nitem > 48 Then
        nPagina = nPagina + 1
        sCad = sCad & Chr(12)
        sCad = sCad & CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(sNumPag)), "", dFecSis) & Chr$(10)
        sCad = sCad & String(nTamMayor + nTamMenor + 2 + nTamAdd, "-") & Chr$(10)
        sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & Chr$(10)
        sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & Chr$(10)
        sCad = sCad & "FECHA DE CANCELACION  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & Space(35) & "HUELLA" & Chr$(10)
        sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & String(nTamAdd, "-") & "+" & Chr$(10)
        nitem = 12
    End If
    
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL2 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL3 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL4 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL5 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    
    sL1 = ""
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & CentrarCadena("........................................", nTamMenor - 1) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & Chr$(10)
    sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-" & "+") & "+" & String(nTamAdd, "-") & "+" & Chr$(10)
         
    nitem = nitem + 12
    rsProdPers.MoveNext
    rsPers.Close
    Set rsPers = Nothing
Loop
Dim arreglocad As String
If nitem > 48 Then
    nPagina = nPagina + 1
    sCad = sCad & Chr(12)
    sCad = sCad & "" 'CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(Str(sNumPag)), "", dFecSis) & Chr$(10)
    arreglocad = CabeRepoCaptac(" ", " ", 1, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(sNumPag)), "", dFecSis)
    sCad = sCad & arreglocad & Chr$(10)
    sCad = sCad & String(nCarLin, "-") & Chr$(10)
    sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & Chr(10)
    sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & Chr$(10)
    sCad = sCad & "FECHA DE CANCELACION  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & Chr(10)
    sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & Chr(10)
    nitem = 12
End If

sL1 = CentrarCadena("INFORME DE ASESORIA LEGAL", nTamMayor)
sL2 = CentrarCadena("FORMAS DE INTERVENCION", nTamMayor)
sL3 = CentrarCadena("[ ]ABRIR   [ ]CERRAR   [ ]RETIRAR", nTamMayor)
sL4 = ""

sCad = sCad & ImpreCarEsp(CentrarCadena(sL1, nTamMayor + nTamMenor + nTamAdd) & " ¦") & Chr$(10)
sCad = sCad & ImpreCarEsp(CentrarCadena(sL2, nTamMayor + nTamMenor + nTamAdd) & " ¦") & Chr$(10)
sCad = sCad & ImpreCarEsp(CentrarCadena(sL3, nTamMayor + nTamMenor + nTamAdd) & " ¦") & Chr$(10)
sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & Chr$(10)
sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & Chr$(10)
sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & Chr$(10)
sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & Chr$(10)
sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & Chr$(10)
sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & Chr$(10)
sCad = sCad & String(nTamMayor, "-") & "-" & String(nTamMenor + nTamAdd, "-") & "-+" & Chr$(10)

sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena("_____________________________               _____________________________             _____________________________", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena("    AUXILIAR DE AHORROS                       ASISTENTE DE OPERACIONES                       ASESORIA LEGAL        ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & Chr$(10)
sCad = sCad & String(nTamMayor, "-") & "-" & String(nTamMenor + nTamAdd, "-") & "-+"
    
GeneraRegistroFirmasCANCELACION = sCad
End Function




Public Sub ActualizaBloqueos(ByVal sCuenta As String, ByVal rsRet As Recordset, _
        ByVal rsTot As Recordset, ByVal sMovNro As String, ByVal nEstado As CaptacEstado)

Dim nEstadoCta As CaptacEstado
Dim clsMant As DCapMantenimiento, clsCap As DCapMovimientos
Dim sComentario As String, sFlag As String, sMov As String
Dim nMotRet As CaptacMotBloqueoRet
Dim nMotTot As CaptacMotBloqueoTot
Dim bIniTran As Boolean, bEstado As Boolean, nMovNro As Long
Dim RSTMP As Recordset
Dim sOperacion As String

Set clsCap = New DCapMovimientos

Set RSTMP = GetDatosCuenta(sCuenta)

On Error GoTo ErrActBloqueo
Set clsMant = New DCapMantenimiento
clsMant.dbCmact.BeginTrans
bIniTran = True
nEstadoCta = gCapEstActiva
If Not rsRet Is Nothing Then
    rsRet.MoveFirst
    Do While Not rsRet.EOF
        nMotRet = CLng(Trim(Right(rsRet("Motivo"), 3)))
        bEstado = IIf(rsRet("Est") = "0", False, True)
        sComentario = rsRet("Comentario")
        sFlag = rsRet("Flag")
        sMov = rsRet("cMovNro")
        
        If sFlag = "N" And bEstado Then
            clsMant.NuevoBloqueoRetiro sCuenta, nMotRet, sComentario, sMovNro
            nEstadoCta = gCapEstBloqRetiro
            
            If nMotRet = gCapMotBlqRetGarantia Or nMotRet = 15 Then
                sOperacion = IIf(Mid(sCuenta, 6, 3) = "232", "201101", IIf(Mid(sCuenta, 6, 3) = "233", "211101", "221101"))
                clsCap.AgregaMov sMovNro, sOperacion, sComentario, gMovEstContabMovContable, gMovFlagVigente
                nMovNro = clsCap.GetnMovNro(sMovNro)
                clsCap.AgregaMovCap nMovNro, sOperacion, sCuenta, RSTMP.Fields("nsaldo"), RSTMP.Fields("nsaldodisp"), RSTMP.Fields("nsaldo")
                clsCap.AgregaMovCapDet nMovNro, sOperacion, sCuenta, gConcCapital, RSTMP.Fields("nsaldo")
            End If
            
        ElseIf sFlag = "A" And Not bEstado Then 'Desbloquea los que ya fueron desmarcados
            clsMant.ActualizaBloqueoRet sCuenta, sComentario, sMovNro, nMotRet, sMov
            
            If nMotRet = gCapMotBlqRetGarantia Or nMotRet = 15 Then
                sOperacion = IIf(Mid(sCuenta, 6, 3) = "232", "201102", IIf(Mid(sCuenta, 6, 3) = "233", "211102", "221102"))
                clsCap.AgregaMov sMovNro, sOperacion, sComentario, gMovEstContabMovContable, gMovFlagVigente
                nMovNro = clsCap.GetnMovNro(sMovNro)
                clsCap.AgregaMovCap nMovNro, sOperacion, sCuenta, RSTMP.Fields("nsaldo"), RSTMP.Fields("nsaldodisp"), RSTMP.Fields("nsaldo")
                clsCap.AgregaMovCapDet nMovNro, sOperacion, sCuenta, gConcCapital, RSTMP.Fields("nsaldo")
            End If
            
        ElseIf sFlag = "A" And bEstado Then
                clsMant.ActualizaComentBlqRet sCuenta, sComentario, nMotRet, sMov
                nEstadoCta = gCapEstBloqRetiro
        End If
        
        
        rsRet.MoveNext
    Loop
End If
If Not rsTot Is Nothing Then
    rsTot.MoveFirst
    Do While Not rsTot.EOF
        nMotTot = CLng(Trim(Right(rsTot("Motivo"), 3)))
        bEstado = IIf(rsTot("Est") = "0", False, True)
        sComentario = rsTot("Comentario")
        sFlag = rsTot("Flag")
        sMov = rsTot("cMovNro")
        If sFlag = "N" And bEstado Then 'Agrega los nuevos bloqueos
            clsMant.NuevoBloqueoTotal sCuenta, nMotTot, sComentario, sMovNro
            nEstadoCta = gCapEstBloqTotal
            
            If nMotTot = gCapMotBlqRetGarantia Or nMotTot = 15 Then
                sOperacion = IIf(Mid(sCuenta, 6, 3) = "232", "201101", IIf(Mid(sCuenta, 6, 3) = "233", "211101", "221101"))
                clsCap.AgregaMov sMovNro, sOperacion, sComentario, gMovEstContabMovContable, gMovFlagVigente
                nMovNro = clsCap.GetnMovNro(sMovNro)
                clsCap.AgregaMovCap nMovNro, sOperacion, sCuenta, RSTMP.Fields("nsaldo"), RSTMP.Fields("nsaldodisp"), RSTMP.Fields("nsaldo")
                clsCap.AgregaMovCapDet nMovNro, sOperacion, sCuenta, gConcCapital, RSTMP.Fields("nsaldo")
            End If
            
        ElseIf sFlag = "A" And Not bEstado Then 'Desbloquea los que ya fueron desmarcados
            clsMant.ActualizaBloqueoTot sCuenta, sComentario, sMovNro, nMotTot, sMov
            
            If nMotTot = gCapMotBlqRetGarantia Or nMotTot = 15 Then
                sOperacion = IIf(Mid(sCuenta, 6, 3) = "232", "201102", IIf(Mid(sCuenta, 6, 3) = "233", "211102", "221102"))
                clsCap.AgregaMov sMovNro, sOperacion, sComentario, gMovEstContabMovContable, gMovFlagVigente
                nMovNro = clsCap.GetnMovNro(sMovNro)
                clsCap.AgregaMovCap nMovNro, sOperacion, sCuenta, RSTMP.Fields("nsaldo"), RSTMP.Fields("nsaldodisp"), RSTMP.Fields("nsaldo")
                clsCap.AgregaMovCapDet nMovNro, sOperacion, sCuenta, gConcCapital, RSTMP.Fields("nsaldo")
            End If
            
        ElseIf sFlag = "A" And bEstado Then
            clsMant.ActualizaComentBlqTot sCuenta, sComentario, nMotTot, sMov
            nEstadoCta = gCapEstBloqTotal
            
            
        End If
        rsTot.MoveNext
    Loop
End If
clsMant.ActualizaEstadoCuenta sCuenta, nEstadoCta
clsMant.dbCmact.CommitTrans
Set clsMant = Nothing
Exit Sub
ErrActBloqueo:
    If bIniTran Then
        clsMant.dbCmact.RollbackTrans
    End If
    Set clsMant = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Function GetCapTasaInteres(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
    ByVal nTipoTasa As CaptacTipoTasa, Optional nPlazo As Double = 0, Optional nValor As Double = 0, _
    Optional sCodage As String = "", Optional bOrdPag As Boolean = False) As Double

Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
GetCapTasaInteres = clsMant.GetCapTasaInteres(nProducto, nMoneda, nTipoTasa, nPlazo, nValor, sCodage, bOrdPag)
Set clsMant = Nothing
End Function

Public Function GetTasaNominal(ByVal nTasaEfecAn As Double, ByVal nPeriodo As Integer) As Double
GetTasaNominal = (((1 + nTasaEfecAn / 100) ^ (1 / nPeriodo)) - 1) * 100 * nPeriodo
End Function

Public Function GetTasaEfectiva(ByVal nTasaNomAn As Double, ByVal nPeriodo As Integer) As Double
'GetTasaEfectiva = (((1 + nTasaNomAn / 36000) ^ nPeriodo) - 1) * 100
GetTasaEfectiva = ((1 + nTasaNomAn / 36000) ^ nPeriodo - 1) * 100

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
Dim bPrimero As Boolean, dFecCan As Date, sFecCan As String
sCad = ""
sNumPag = ""
nLinPag = 0
nCarLin = 85
sMoneda = IIf(nMoneda = gMonedaNacional, "SOLES", "DOLARES")
sTitRp1 = "CRONOGRAMA  DE  RETIRO  DE INTERESES"
sTitRp2 = ""
sCodCta = ""
nCntPag = 1
nAcum = 0
dAuxiliar = dFecApe
dFecVenc = DateAdd("d", nPlazo, dFecApe)
dFecCan = DateAdd("d", nPlazo + 1, dFecApe)
sFecVenc = Format$(dFecVenc, "dd mmmm yyyy")
sFecCan = Format$(dFecCan, "dd mmmm yyyy")

sNumPag = FillNum(Trim(Str(nCntPag)), 4, " ")
sCad = sCad & CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, sNumPag, "", dFecVenc) & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & String(3, " ") & "PLAZO             : " & nPlazo & Space(3) & " DIAS" & Chr$(10)
sCad = sCad & String(3, " ") & "CAPITAL           : " & Format$(nCapital, "#,##0.00") & Space(3) & sMoneda & Chr$(10)
sCad = sCad & String(3, " ") & "TASA              : " & nTasa & " %" & Chr$(10)
If nIntPriMes > 0 Then
    sCad = sCad & String(3, " ") & "FECHA APE. CHQ.   : " & Format$(dFecApe, "dd mmmm yyyy") & Chr$(10)
    sCad = sCad & String(3, " ") & "FECHA VAL. CHQ.   : " & Format$(dFecVal, "dd mmmm yyyy") & Chr$(10)
Else
    sCad = sCad & String(3, " ") & "FECHA APERTURA    : " & Format$(dFecApe, "dd mmmm yyyy") & Chr$(10)
End If
sCad = sCad & String(3, " ") & "FECHA VENCIMIENTO : " & sFecVenc & Chr$(10)
sCad = sCad & String(3, " ") & "FECHA CANCELACION : " & sFecCan & Chr$(10)
sCad = sCad & String(3, " ") & "MONTO FINAL PLAZO : " & Format$(nIntFinPlazo, "#,##0.00") & " (Sin retirar intereses mensualmente)" & Chr$(10)
sCad = sCad & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & "  FECHA OPERACION            IMPORTE   " & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & Chr$(10)
nLinPag = 11
dAuxiliar = DateAdd("d", 30, dAuxiliar)
bPrimero = True
While dFecVenc > dAuxiliar
    sFchtra = Format(dAuxiliar, "dd/mm/yyyy")
    If bPrimero Then
        If nIntPriMes > 0 Then
            nAcum = nAcum + nIntPriMes
            sCad = sCad & "   " & sFchtra & "           " & JDNum(Trim(Str(nIntPriMes)), 12, True, 9, 2) & Chr$(10)
        Else
            nAcum = nAcum + nInteres
            sCad = sCad & "   " & sFchtra & "           " & JDNum(Trim(Str(nInteres)), 12, True, 9, 2) & Chr$(10)
        End If
        bPrimero = False
    Else
        sCad = sCad & "   " & sFchtra & "           " & JDNum(Trim(Str(nInteres)), 12, True, 9, 2) & Chr$(10)
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
    sCad = sCad & "   " & Format(dFecVenc, "dd/mm/yyyy") & "           " & JDNum(Trim(Str(nUltInteres)), 12, True, 9, 2) & "    RENOVACION" & Chr$(10)
Else
    nAcum = nAcum + nInteres
    sCad = sCad & "   " & Format(dFecVenc, "dd/mm/yyyy") & "           " & JDNum(Trim(Str(nInteres)), 12, True, 9, 2) & "    RENOVACION" & Chr$(10)
End If
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & String(3, " ") & "ACUMULADO " & FillText(sMoneda, 8, " ") & "   " & JDNum(Trim(Str(nAcum)), 12, True, 9, 2) & Chr$(10)
sCad = sCad & String(nCarLin, "=") & Chr$(10)
GetPFPlanRetInt = sCad
End Function

Public Function GetBeneficiarios(ByVal sPersona As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetBeneficiarios = clsMant.GetBeneficiarios(sPersona)
Set clsMant = Nothing
End Function

Public Function GetTarjetaCuentas(ByVal sTarjeta As String, Optional nProducto As Producto = gCapTodoAhorro) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetTarjetaCuentas = clsMant.GetTarjetaCuentas(sTarjeta, nProducto)
Set clsMant = Nothing
End Function

Public Function GetTarjetaEstadoHist(ByVal sTarjeta As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetTarjetaEstadoHist = clsMant.GetTarjetaEstadoHist(sTarjeta)
Set clsMant = Nothing
End Function

Public Function GetListadoCuentasCTS(ByVal sPerscod As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetListadoCuentasCTS = clsMant.GetListadoCuentasCTS(sPerscod)
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

    If bTrans And Err.Number <> 91 Then
        clsMant.dbCmact.RollbackTrans
    End If
'    If bTrans And Err.Number = 91 Then
'        clsmant.dbCmact.CommitTrans
'    End If
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

Public Function GetPersonaCuenta(ByVal sCuenta As String, Optional nTipRel As CaptacRelacPersona) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetPersonaCuenta = clsMant.GetPersonaCuenta(sCuenta, nTipRel)
Set clsMant = Nothing
End Function

Public Function GetPersonasCuentaNA(ByVal sCuenta As String, ByVal sPerscod As String, Optional ByVal nPersoneria As Integer) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetPersonasCuentaNA = clsMant.GetPersonasCuentaNA(sCuenta, sPerscod, nPersoneria)
Set clsMant = Nothing
End Function


Public Function GeneraConvenioTarjeta(ByVal sCuenta As String, ByVal dFecha As Date) As String
Dim rsPers As Recordset
Dim lsTit As String, lsRep As String
Dim lsDireccion As String, lsDNI As String
Dim sCad As String, sPersona As String, saux As String
Dim clsMant As DCapMantenimiento

sCad = Chr$(27) & Chr$(77)
Set clsMant = New DCapMantenimiento
Set rsPers = clsMant.GetPersonaCuenta(sCuenta)
Set clsMant = Nothing

Open App.path & "\FormatoCarta\ConTar.txt" For Input As #1

Do While Not EOF(1)
    Line Input #1, saux
    sCad = sCad & Chr$(10) & saux
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
                    lsRep = lsRep & Chr(10) & Space(40) & PstaNombre(Trim(rsPers("Nombre")))
                End If
            Else
                If lsTit = "" Then
                    lsDNI = rsPers("ID") & " " & rsPers("ID N°")
                    lsTit = PstaNombre(rsPers("Nombre"))
                    lsDireccion = rsPers("Direccion") & ""
                Else
                    lsTit = lsTit & Chr(10) & Space(52) & PstaNombre(rsPers("Nombre"))
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



Public Function GetCartasOP(ByVal sAge As String) As Recordset
'Dim rsOP As Recordset
Dim lsTit As String, lsRep As String
Dim lsDireccion As String, lsDNI As String
Dim sCad As String, sPersona As String, saux As String
Dim clsMant As DCapMantenimiento
Dim LALISTA() As String
Dim lnFilas As Long, lnItem As Long
Dim lsCartaModelo As String
Dim lsCadImp As String

   Set clsMant = New DCapMantenimiento

   Set GetCartasOP = clsMant.GetDatosOP(sAge)
   Set clsMant = Nothing
     
  '  Set rsOP = clsMant.GetDatosOP(sAge)
    
    
  
'    lnFilas = 0
'    ReDim LALISTA(rsOP.RecordCount, 5)
'
'    While Not rsOP.EOF
'        lnFilas = lnFilas + 1
'                    LALISTA(lnFilas, 0) = rsOP!Nombre
'                    LALISTA(lnFilas, 1) = rsOP!Direccion
'                    LALISTA(lnFilas, 2) = rsOP!CCTACODANT
'                    LALISTA(lnFilas, 3) = rsOP!cCtaCod
'
'        rsOP.MoveNext
'     Wend
     
'      For lnItem = 1 To lnFilas
'        lsCartaModelo = psTextoCarta
'        lsCartaModelo = Replace(lsCartaModelo, "<<DIA>>", Format(dFecha, "dd"))
'        lsCartaModelo = Replace(lsCartaModelo, "<<MES>>", Format(dFecha, "MM"))
'        lsCartaModelo = Replace(lsCartaModelo, "<<ANIO>>", Format(dFecha, "YYYY"))
'        lsCartaModelo = Replace(lsCartaModelo, "<<NOMBRE>>", PstaNombre(LALISTA(lnItem, 0), False), , 1, vbTextCompare)
'        lsCartaModelo = Replace(lsCartaModelo, "<<DIRECCION>>", LALISTA(lnItem, 1), , 1, vbTextCompare)
'        lsCartaModelo = Replace(lsCartaModelo, "<<CODIGO ANTIGUO>>", LALISTA(lnItem, 2), , 1, vbTextCompare)
'        lsCartaModelo = Replace(lsCartaModelo, "<<CODIGO NUEVO>>", LALISTA(lnItem, 3), , 1, vbTextCompare)
'
'
'        lsCadImp = lsCadImp & lsCartaModelo & Chr(12)
'        'prgList.Value = vCont
'    Next lnItem




    
'    Erase LALISTA
    'nRemImprimeAvisoRemate = lsCadImp
          
'Set clsMant = Nothing
'        rsOP.Close
'        Set rsOP = Nothing
'        GeneraCartasOP = lsCadImp
End Function


Public Function GetNombreTitulares(ByVal sCuenta As String, Optional pbNombreCompleto As Boolean = False, Optional pnTitulares As Integer = 0, Optional pnMargenIzquierdo As Integer = 0) As String
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
GetNombreTitulares = clsMant.GetNombreTitulares(sCuenta, pbNombreCompleto, pnTitulares, pnMargenIzquierdo)
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
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    If nErrNum = 1 Then
        MsgBox "Rango ya emitido. Rango Referencia: " & Chr$(13) & sRango, vbInformation, "Aviso"
    Else
        MsgBox Err.Description, vbInformation, "Aviso"
    End If
End Sub

Public Sub AnulaRangoOrdenPago(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nNumIni As Long, ByVal nNumFin As Long, ByVal sGlosa As String)

Dim clsMant As DCapMantenimiento
Dim nErrNum As Integer, nAnuladas As Integer
Dim i As Long
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
For i = nNumIni To nNumFin
    If clsMant.ExisteOrdenPagoEmitidas(sCuenta, i, i) = "" Then
        sMsgErr = sMsgErr & "Orden Pago NO Emitida : " & i & Chr$(13)
    Else
        Set rsEstado = clsMant.GetEstadoOrdenPagoEmitida(sCuenta, i)
        If rsEstado.EOF And rsEstado.BOF Then
            nEstado = gCapOPEstEmitida
            sEstado = ""
        Else
            nEstado = rsEstado("nEstado")
            sEstado = rsEstado("Estado")
        End If
        If nEstado <> gCapOPEstEmitida And nEstado <> gCapOPEstExtornada Then
            sMsgErr = sMsgErr & "Orden Pago NO Válida : " & i & Space(2) & sEstado & Chr$(13)
        Else
            If nEstado = gCapOPEstEmitida Then clsMant.AgregaOrdenPagoRecibida sCuenta, i, 0
            clsMant.AnulaOrdenPagoEmitida sCuenta, i, sMovNro, 0
            nAnuladas = nAnuladas + 1
        End If
    End If
Next i
If nAnuladas = 0 Then
    clsMant.dbCmact.RollbackTrans
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
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    If nErrNum = 0 Then
        MsgBox Err.Description, vbInformation, "Aviso"
    ElseIf nErrNum = 1 Then
        MsgBox "Ordenes Anuladas : " & nAnuladas & Chr$(13) & sMsgErr, vbInformation, "Aviso"
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
    nEstado = rsEstado("nEstado")
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
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    If nErrNum = 0 Then
        MsgBox Err.Description, vbInformation, "Aviso"
    ElseIf nErrNum = 1 Then
        MsgBox "Orden Pago NO Emitida", vbInformation, "Aviso"
    ElseIf nErrNum = 2 Then
        MsgBox "Orden Pago NO Válida. Estado " & sEstado, vbInformation, "Aviso"
    End If
End Sub

Public Function GetSaldoFecha(ByVal sCuenta As String, ByVal dFecha As Date) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetSaldoFecha = clsMant.GetSaldoFecha(sCuenta, dFecha)
Set clsMant = Nothing
End Function
Public Function GetSaldoFecha2(ByVal cUser As String, ByVal dFecFin As Date) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetSaldoFecha2 = clsMant.GetSaldoFecha2(cUser, dFecFin)
Set clsMant = Nothing
End Function

Public Function GetMovimientosCuenta(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, Optional bMuestraExtornos As Boolean = True) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetMovimientosCuenta = clsMant.GetMovimientosCuenta(sCuenta, dFecIni, dFecFin, nNumMov, bMuestraExtornos)
Set clsMant = Nothing
End Function

Public Function GetMovimientosCuenta2(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, Optional bMuestraExtornos As Boolean = True, Optional cUser As String, Optional ByRef RSTEMP As Recordset) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetMovimientosCuenta2 = clsMant.GetMovimientosCuenta2(sCuenta, dFecIni, dFecFin, nNumMov, bMuestraExtornos, cUser, RSTEMP)
Set clsMant = Nothing
End Function

Public Function GetMovimientosCuenta4(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, Optional bMuestraExtornos As Boolean = True, Optional cUser As String, Optional ByRef RSTEMP As Recordset) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetMovimientosCuenta4 = clsMant.GetMovimientosCuenta4(sCuenta, dFecIni, dFecFin, nNumMov, bMuestraExtornos, cUser, RSTEMP)
Set clsMant = Nothing
End Function


Public Sub ActualizaSobregiro(ByVal sCuenta As String, ByVal nSobregiro As Long, _
        ByVal sMovNro As String, ByVal nEstadoCta As CaptacEstado, Optional bBloqueaCuenta As Boolean = False)
Dim clsMant As DCapMantenimiento
Dim bTrans As Boolean

Set clsMant = New DCapMantenimiento

On Error GoTo ErrGraba
clsMant.dbCmact.BeginTrans
bTrans = True
If bBloqueaCuenta Then
    clsMant.NuevoBloqueoRetiro sCuenta, gCapMotBlqRetOrdenPagoRechazada, "Orden Sobregirada", sMovNro
    If nEstadoCta = gCapEstActiva Then
        clsMant.ActualizaEstadoCuenta sCuenta, gCapEstBloqRetiro
        clsMant.AgregaCaptacEstado sCuenta, gCapEstBloqRetiro, sMovNro
    End If
End If
clsMant.ActualizaSobregiro sCuenta, nSobregiro
clsMant.dbCmact.CommitTrans
bTrans = False
Set clsMant = Nothing
Exit Sub
ErrGraba:
    If bTrans Then
        clsMant.dbCmact.RollbackTrans
    End If
    Set clsMant = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Function GetTarifaOrdenPago(ByVal nMoneda As Moneda) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetTarifaOrdenPago = clsMant.GetTarifaOrdenPago(nMoneda)
Set clsMant = Nothing
End Function

Public Function GetMaxOrdPagEmitida(ByVal sCuenta As String) As Long
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
GetMaxOrdPagEmitida = clsMant.GetMaxOrdPagEmitida(sCuenta)
Set clsMant = Nothing
End Function

Public Function GetHistOrdPagEmision(ByVal sCuenta As String) As Recordset
Dim clsMant As DCapMantenimiento
Set clsMant = New DCapMantenimiento
Set GetHistOrdPagEmision = clsMant.GetHistOrdPagEmision(sCuenta)
Set clsMant = Nothing
End Function

Public Sub AgregaSolicitudOrdPagTal(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nInicio As Long, ByVal nFin As Long, ByVal nNumTal As Long)
Dim oCapMant As DCapMantenimiento

Dim bTrans As Boolean
Set oCapMant = New DCapMantenimiento
bTrans = True
oCapMant.AgregaCapOrdPagEmision sCuenta, sMovNro, gCapTalOrdPagEstSolicitado, nInicio, nFin, nNumTal
oCapMant.AgregaCapOrdPagEstado sCuenta, sMovNro, gCapTalOrdPagEstSolicitado, nInicio
bTrans = False
Set oCapMant = Nothing
Exit Sub
ErrGraba:
    If bTrans Then oCapMant.dbCmact.RollbackTrans
    Set oCapMant = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Function GetCapOrdPagEmision(ByVal nEstado As CapOrdPagTalEstado) As Recordset
Dim oCapMant As DCapMantenimiento
Set oCapMant = New DCapMantenimiento
Set GetCapOrdPagEmision = oCapMant.GetCapOrdPagEmision(nEstado)
Set oCapMant = Nothing
End Function

Public Sub ActualizaCapOrdPagEmision(ByVal sCuenta As String, ByVal sMovNroNuevo As String, _
        ByVal nEstadoNuevo As CapOrdPagTalEstado, ByVal sMovNroAnt As String, _
        ByVal nEstadoAnt As CapOrdPagTalEstado, ByVal nInicio As Long)

Dim oCap As DCapMantenimiento
Set oCap = New DCapMantenimiento
oCap.dbCmact.BeginTrans
oCap.ActualizaCapOrdPagEmision sCuenta, sMovNroNuevo, nEstadoNuevo, sMovNroAnt, nEstadoAnt
oCap.AgregaCapOrdPagEstado sCuenta, sMovNroNuevo, nEstadoNuevo, nInicio
oCap.dbCmact.CommitTrans
End Sub

Private Function GetTitulares(ByVal sCuenta As String) As String
Dim sNombre1 As String * 60, sNombre2 As String * 60, sNombre3 As String * 60
Dim sDocId1 As String * 11, sDocId2 As String * 11, sDocId3 As String * 11
Dim sNom As String * 60, sDoc As String * 11
Dim rsPers As Recordset
Dim oCapMant As NCapMantenimiento
Dim i As Integer

sNombre1 = Space(60)
sNombre2 = Space(60)
sNombre3 = Space(60)
sDocId1 = Space(11)
sDocId2 = Space(11)
sDocId3 = Space(11)
i = 0

Set oCapMant = New NCapMantenimiento
Set rsPers = oCapMant.GetPersonaCuenta(sCuenta)
If Not (rsPers.EOF And rsPers.BOF) Then
    Do While Not rsPers.EOF
        sNom = PstaNombre(rsPers("Nombre"), True)
        sDoc = rsPers("ID N°")
        i = i + 1
        Select Case i
            Case 1
                sNombre1 = sNom
                sDocId1 = sDoc
            Case 2
                sNombre2 = sNom
                sDocId2 = sDoc
            Case 3
                sNombre3 = sNom
                sDocId3 = sDoc
        End Select
        rsPers.MoveNext
    Loop
    GetTitulares = sDocId1 & sDocId2 & sDocId3 & sNombre1 & sNombre2 & sNombre3
Else
    GetTitulares = ""
End If
rsPers.Close
Set rsPers = Nothing
End Function

Private Function GetMiembrosCta(ByVal sCuenta As String) As String
Dim sNombre1 As String * 60, sNombre2 As String * 60, sNombre3 As String * 40
Dim sDocId1 As String * 11, sDocId2 As String * 11, sDocId3 As String * 11
Dim sNom As String * 60, sDoc As String * 11
Dim sNom1 As String * 40
Dim rsPers As Recordset
Dim oCapMant As NCapMantenimiento
Dim i As Integer

sNombre1 = Space(60)
sNombre2 = Space(60)
'****POR MODIFICACION
'sNombre3 = Space(60)
sNombre3 = Space(40)
sDocId1 = Space(11)
sDocId2 = Space(11)
sDocId3 = Space(11)
i = 0

Set oCapMant = New NCapMantenimiento

LblSeguir:

If i = 0 Then
    Set rsPers = oCapMant.GetPersonaCuenta(sCuenta, gCapRelPersTitular)
Else
    Set rsPers = oCapMant.GetPersonaCuenta(sCuenta, gCapRelPersRepTitular)
    
End If

If Not (rsPers.EOF And rsPers.BOF) Then
    Do While Not rsPers.EOF
        sNom = PstaNombre(rsPers("Nombre"), True)
        sNom1 = PstaNombre(rsPers("Nombre"), True)
        sDoc = rsPers("ID N°")
        
'        If i = 0 Then
'            If rsPers("nperspersoneria") <> 1 Then
'                sNombre1 = IIf(Len(Trim(sNom)) < 60, Trim(sNom) + Space(60 - Len(Trim(sNom))), Left(Trim(sNom), 60))
'                sDocId1 = IIf(Len(Trim(sDoc)) < 11, Trim(sDoc) + Space(11 - Len(Trim(sDoc))), Left(Trim(sDoc), 11))
'                i = i + 1
'               GoTo LblSeguir
'            End If
'        End If
      
        Select Case i
            Case 0
                sNombre1 = IIf(Len(Trim(sNom)) < 60, Trim(sNom) + Space(60 - Len(Trim(sNom))), Left(Trim(sNom), 60))
                sDocId1 = IIf(Len(Trim(sDoc)) < 11, Trim(sDoc) + Space(11 - Len(Trim(sDoc))), Left(Trim(sDoc), 11))
            Case 1
                sNombre2 = IIf(Len(Trim(sNom)) < 60, Trim(sNom) + Space(60 - Len(Trim(sNom))), Left(Trim(sNom), 60))
                sDocId2 = IIf(Len(Trim(sDoc)) < 11, Trim(sDoc) + Space(11 - Len(Trim(sDoc))), Left(Trim(sDoc), 11))
            Case 2
                'sNombre3 = IIf(Len(Trim(sNom)) < 60, Trim(sNom) + Space(60 - Len(Trim(sNom))), Left(Trim(sNom), 60))
                sNombre3 = IIf(Len(Trim(sNom1)) < 40, Trim(sNom1) + Space(40 - Len(Trim(sNom1))), Left(Trim(sNom1), 40))
                sDocId3 = IIf(Len(Trim(sDoc)) < 11, Trim(sDoc) + Space(11 - Len(Trim(sDoc))), Left(Trim(sDoc), 11))
            Case Else
                Exit Do
        End Select
          i = i + 1
        rsPers.MoveNext
    Loop
    
    'GetTitulares = sDocId1 & sDocId2 & sDocId3 & sNombre1 & sNombre2 & sNombre3
    GetMiembrosCta = sNombre1 & sNombre2 & sNombre3 & sDocId1 & sDocId2 & sDocId3
Else
    GetMiembrosCta = ""
End If
rsPers.Close
Set rsPers = Nothing
End Function


Public Function GeneraArchivoEnvioOrdPagTal(ByVal dFecSis As Date, ByVal rsOrd As Recordset, _
            ByVal sMovNro As String) As String
Dim sRuta As String
Dim oCap As DCapMantenimiento
Dim bTrans As Boolean
Dim sCuenta As String
Dim sInicio As String * 8, sFin As String * 8
Dim sNumTal As String * 2
Dim sOficCentral As String * 3
Dim sAgencia As String * 3
Dim sTitular As String
Dim sCadena As String, sFecha As String
Dim sMoneda As String * 1
Dim sCadImp As String
Dim sOPTipoTalon  As String
Dim sMovNroAnt As String
'sRuta = App.path & "\SPOOLER\" & "CMACICA-" & Right$(Year(gdFecSis), 1) & Format$(Month(gdFecSis), "00") & Format$(Day(gdFecSis), "00") & ".TXT"

'sRuta = App.path & "\SPOOLER\" & "CICA" & IIf(Len(Day(dFecSis)) < 2, "0", "") & CStr(Day(dFecSis)) & IIf(Len(Month(dFecSis) < 2), "0", "") & CStr(Month(dFecSis)) & CStr(Year(dFecSis)) & Left(GetHoraServer(), 2) & Mid(GetHoraServer(), 4, 2) & Right(GetHoraServer(), 2) & ".TXT"
sRuta = App.path & "\SPOOLER\" & "CICA" & IIf(Len(Day(dFecSis)) < 2, "0", "") & CStr(Day(dFecSis)) & IIf(Len(Month(dFecSis) < 2), "0", "") & CStr(Month(dFecSis)) & ".TXT"


Dim fs As New Scripting.FileSystemObject
Set fs = New Scripting.FileSystemObject
If fs.FileExists(sRuta) Then
    If MsgBox("Ya existe un archivo generado el dia de hoy.." & vbCrLf & "¿Desea BORRAR este archivo y GENERAR uno nuevo?", vbYesNo + vbDefaultButton2 + vbQuestion, "AVISO") = vbYes Then
      fs.DeleteFile sRuta, True
    End If
End If


Open sRuta For Output As #1

Set oCap = New DCapMantenimiento
oCap.dbCmact.BeginTrans
bTrans = True
On Error GoTo ErrGraba
Do While Not rsOrd.EOF
    If rsOrd("Act") = "1" Then
    
        sCuenta = rsOrd("Cuenta")
        sInicio = rsOrd("Inicio")
        sFin = rsOrd("Fin")
        sNumTal = Format$(rsOrd("# Tal"), "00")
        sOPTipoTalon = rsOrd("nTipo")
        sMovNroAnt = rsOrd("cMovNro")
        sAgencia = Format$(CLng(Mid$(sCuenta, 4, 2)), "000")
        sOficCentral = "001"
        sMoneda = Mid$(sCuenta, 9, 1)
        sTitular = GetMiembrosCta(sCuenta)
        If sTitular <> "" Then
            'sCadena = sInicio & Mid$(sCuenta, 9, 10) & sTitular & sAgencia & sOficCentral & sNumTal & "0" & sOPTipoTalon & sMoneda
            sCadena = sInicio & Mid$(sCuenta, 9, 10) & sTitular & sAgencia & sAgencia & sNumTal & "1" & sOPTipoTalon & sMoneda & sAgencia
            sCadImp = sCadImp & sCadena
            Print #1, sCadena
            oCap.ActualizaCapOrdPagEmision sCuenta, sMovNro, gCapTalOrdPagEstEnviado, sMovNroAnt, gCapTalOrdPagEstSolicitado
            oCap.AgregaCapOrdPagEstado sCuenta, sMovNro, gCapTalOrdPagEstEnviado, CLng(sInicio)
        Else
            MsgBox "No se encontraron titulares para la cuenta " & sCuenta, vbInformation, "Aviso"
        End If
        
        
    End If
    rsOrd.MoveNext
Loop

Close #1
oCap.dbCmact.CommitTrans
bTrans = False
GeneraArchivoEnvioOrdPagTal = sRuta

Exit Function
ErrGraba:
    If bTrans Then oCap.dbCmact.RollbackTrans
    Set oCap = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
    GeneraArchivoEnvioOrdPagTal = ""
End Function

Public Sub ActualizaCapOrdPagEstado(ByVal dFecSis As Date, ByVal rsOrd As Recordset, _
            ByVal sMovNro As String, ByVal nEstadoNuevo As CapOrdPagTalEstado, _
            ByVal nEstadoAnt As CapOrdPagTalEstado)

Dim oCap As DCapMantenimiento

Dim sCuenta As String
Dim sMovNroAnt As String
Dim nInicio As Long
Dim bTrans As Boolean

Set oCap = New DCapMantenimiento
oCap.dbCmact.BeginTrans
bTrans = True
On Error GoTo ErrGraba


Do While Not rsOrd.EOF
    If rsOrd("Act") = "1" Then
        sCuenta = rsOrd("Cuenta")
        sMovNroAnt = rsOrd("cMovNro")
        nInicio = rsOrd("Inicio")
        oCap.ActualizaCapOrdPagEmision sCuenta, sMovNro, nEstadoNuevo, sMovNroAnt, nEstadoAnt
        oCap.AgregaCapOrdPagEstado sCuenta, sMovNro, nEstadoNuevo, nInicio
    End If
    rsOrd.MoveNext
Loop


oCap.dbCmact.CommitTrans
bTrans = False
Exit Sub
ErrGraba:
    If bTrans Then oCap.dbCmact.RollbackTrans
    Set oCap = Nothing
    MsgBox Err.Description, vbExclamation, "Error"
End Sub
Public Sub ImprimeConevenioOP(ByVal rsOrd As Recordset)
Dim oCap As DCapMantenimiento
Dim sCuenta As String
Set oCap = New DCapMantenimiento
EliminaConvenioOP
rsOrd.MoveFirst
Do While Not rsOrd.EOF
    If rsOrd("Act") = "1" Then
        sCuenta = rsOrd("Cuenta")
        GeneraConvenioOP sCuenta  'ppoa  carga temporal
    End If
    rsOrd.MoveNext
Loop

'**************ejrs **********************
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim sql  As String
Set oCon = New DConecta
Set rs = New ADODB.Recordset

MsgBox "Coloque papel para el convenio de Orden de Pago", vbInformation, "Aviso"
oCon.AbreConexion
sql = "SELECT * FROM tmpConvenioOP  "
Set rs = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing

With DRContratoOP
   Set .DataSource = rs
      .DataMember = ""
      .WindowState = vbMaximized
      .Title = "CONVENIO ORDEN DE PAGO"
      .inicio
      .Refresh
      .Show 1
 End With
'DRContratoOP.Show 1
End Sub
Public Function GetCuentaAbonoIF(ByVal sPerscod As String, ByVal nMoneda As Moneda) As String
Dim oCap As DCapMantenimiento
Set oCap = New DCapMantenimiento
GetCuentaAbonoIF = oCap.GetCuentaAbonoIF(sPerscod, nMoneda)
Set oCap = Nothing
End Function

Public Function BuscaCreditosPendientesPago(ByVal sCuenta As String) As Boolean
Dim oCap As DCapMantenimiento
Set oCap = New DCapMantenimiento
BuscaCreditosPendientesPago = oCap.BuscaCreditosPendientesPago(sCuenta)
Set oCap = Nothing
End Function

Public Function GetChequesOperaciones(Optional sCuenta As String = "", Optional sNumCheque As String = "", Optional ByVal nEstado As Integer = -1) As Recordset
Dim oCap As DCapMantenimiento
Set oCap = New DCapMantenimiento
Set GetChequesOperaciones = oCap.GetChequesOperaciones(sCuenta, sNumCheque, nEstado)
Set oCap = Nothing
End Function

Public Function GetChequesRegistrados(ByVal sNumCheque As String) As Recordset
Dim oCap As DCapMantenimiento
Set oCap = New DCapMantenimiento
Set GetChequesRegistrados = oCap.GetChequesRegistrados(sNumCheque)
Set oCap = Nothing
End Function

Public Function GetChequesValorizadosInmediato(ByVal dFecha As Date, Optional sCuenta As String = "", _
            Optional sNumCheque As String = "") As Recordset
Dim oCap As DCapMantenimiento
Set oCap = New DCapMantenimiento
Set GetChequesValorizadosInmediato = oCap.GetChequesValorizadosInmediato(dFecha, sCuenta, sNumCheque)
Set oCap = Nothing
End Function

Public Function GetChequesEnValorizados(sCuenta) As Recordset
Dim oCap As DCapMantenimiento
Set oCap = New DCapMantenimiento
Set GetChequesEnValorizados = oCap.GetChequesEnValorizacion(sCuenta)
Set oCap = Nothing
End Function

Public Function GetCreditosPendientes(ByVal sPersona As String) As Recordset
'Dim rsCred As Recordset
'Dim rsCta As Recordset
'Dim nProducto As Producto
'Set rsCred = New Recordset
'
'Dim oMant As NCapMantenimiento
'Set oMant = New NCapMantenimiento
'Set rsCta = oMant.GetProductoPersonaConsol(sPersona)
'Set oMant = Nothing
'
'Do While Not rsCta.EOF
'
'Loop
'GeneraRecordsetCred rsCred
'rsCred.Open
'CreditosPendientes sPersona, rsCred
'Set GetCreditosPendientes = rsCred
End Function

Public Function CreditosPendientes(ByVal sPersona As String, ByRef rsCred As ADODB.Recordset) As Boolean
'Dim rsCta As ADODB.Recordset
'Dim nAgencias As Long, nLen As Long, nCuentasCre As Long, nCuentasPrd As Long, nCuentasJud As Long
'Dim sBuffer As String, sAgencia As String, sTipo As String
'Dim sSimbMoneda As String, sCuentas As String, sVencidos As String
'Dim I As Integer, J As Integer, nDiasAtraso As Integer
'Dim nDeudaPrend As Double, nMinimoPagar As Double
'Dim bVencidos As Boolean
'Dim nValor1 As Double, nValor2 As Double
'sAgencia = ""
'sBuffer = ""
'nCuentasCre = 0
'
'GetParametrosValores gsBusqCredPendVenc, nValor1, nValor2
'If nValor1 = 1 Then
'    bVencidos = True
'    nDiasAtraso = Int(nValor2)
'Else
'    bVencidos = False
'End If
'Set rsCta = New ADODB.Recordset
'rsCta.CursorLocation = adUseClient
'For I = 1 To UBound(CuentasPersonaCre, 1)
'    sCuentas = "("
'    sAgencia = Left(CuentasPersonaCre(I), 2)
'    For J = I To UBound(CuentasPersonaCre, 1)
'        sBuffer = Left(CuentasPersonaCre(J), 2)
'        If sBuffer = sAgencia Then
'            sCuentas = sCuentas & "'" & CuentasPersonaCre(J) & "',"
'        Else
'            Exit For
'        End If
'    Next J
'    sCuentas = Left(sCuentas, Len(sCuentas) - 1)
'    sCuentas = sCuentas & ")"
'    I = J - 1
'    If AbreConeccion(sAgencia & "0000000000") Then
'        'Verifica primero los creditos pymes pendientes de pago
'        If bVencidos Then
'            sVencidos = "nDiasAtraso > " & nDiasAtraso & " And "
'        Else
'            sVencidos = ""
'        End If
'        VSQL = "Select cCodCta, nDiasAtraso, cRefinan, cNota1, Convert(Varchar(10),dFecVenc,103) dFecVenc, Count(*) nNroCuo, " _
'            & "SUM(nCuota) nCuota, SUM(nMora) nMora, SUM(nGastos) nGastos From ( " _
'            & "Select cCodCta, nDiasAtraso, cRefinan, cNota1, cNroCuo, dFecVenc, nCuota, nMora, SUM(nGastos) nGastos From ( " _
'            & "Select C.cCodCta, C.nDiasAtraso, C.cRefinan, C.cNota1, P.cNroCuo, nCuota = P.nCapital-P.nCapPag+P.nInteres-P.nIntComPag+P.nIntGra+P.nIntGraPag, " _
'            & "nMora = ISNULL(P.nMora-P.nMoraPag,0), nGastos = ISNULL(PG.nMonNeg-PG.nMonPag,0), dFecVenc = (Select Min(dFecVenc) " _
'            & "From PlandesPag P1 Where P1.cCodCta = P.cCodCta And P1.cEstado = 'P' And P1.cTipo = 'C' And DateDiff(dd,P1.dFecVenc,'" & Format$(Date, "mm/dd/yyyy") & "') > 0) " _
'            & "From Credito C INNER JOIN PlandesPag P LEFT JOIN PlanGastos PG ON " _
'            & "P.cCodCta = PG.cCodCta And Convert(Int,P.cNroCuo) = PG.nNroCuo ON C.cCodCta = P.cCodCta " _
'            & "Where " & sVencidos & " C.cEstado = 'F' And P.cEstado = 'P' And P.cTipo = 'C' And " _
'            & "((PG.cAplicado = 'C' And PG.cEstado = 'P') Or (PG.nMonNeg-PG.nMonPag) IS NULL) " _
'            & "And C.cCodCta IN " & sCuentas & " And DateDiff(dd,P.dFecVenc,'" & Format$(Date, "mm/dd/yyyy") & "') > 0 " _
'            & ") T Group by cCodCta, nDiasAtraso, cRefinan, cNota1, cNroCuo, dFecVenc, nCuota, nMora " _
'            & ") T1 Group by cCodCta, nDiasAtraso, cRefinan, cNota1, Convert(Varchar(10),dFecVenc,103)"
'
'        rsCta.Open VSQL, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
'        Set rsCta.ActiveConnection = Nothing
'        If Not (rsCta.EOF And rsCta.BOF) Then
'            Do While Not rsCta.EOF
'                nCuentasCre = nCuentasCre + 1
'                rsCred.AddNew
'                rsCred("Cuenta") = Trim(rsCta("cCodCta"))
'                rsCred("Estado") = IIf(rsCta("cRefinan") = "R", "RF", "")
'                rsCred("Vencimiento") = CDate(rsCta("dFecVenc"))
'                rsCred("CuoVenc") = Trim(rsCta("nNroCuo"))
'                rsCred("Cuota") = rsCta("nCuota")
'                rsCred("Mora") = rsCta("nMora")
'                rsCred("Gastos") = rsCta("nGastos")
'                rsCred("MinPagar") = rsCta("nCuota") + rsCta("nGastos") + rsCta("nMora")
'                rsCred.Update
'                rsCta.MoveNext
'            Loop
'        End If
'        rsCta.Close
'    End If
'Next I
'
''Verifica los Créditos Prendarios
'sAgencia = ""
'sBuffer = ""
'nCuentasPrd = 0
'For I = 1 To UBound(CuentasPersonaPrd, 1)
'    sCuentas = "("
'    sAgencia = Left(CuentasPersonaPrd(I), 2)
'    For J = I To UBound(CuentasPersonaPrd, 1)
'        sBuffer = Left(CuentasPersonaPrd(J), 2)
'        If sBuffer = sAgencia Then
'            sCuentas = sCuentas & "'" & CuentasPersonaPrd(J) & "',"
'        Else
'            Exit For
'        End If
'    Next J
'    sCuentas = Left(sCuentas, Len(sCuentas) - 1)
'    sCuentas = sCuentas & ")"
'    I = J - 1
'    If AbreConeccion(sAgencia & "0000000000") Then
'        If bVencidos Then
'            sVencidos = "And DateDiff(dd,C.dFecVenc,'" & Format$(Date, "mm/dd/yyyy") & "') > " & nDiasAtraso
'        Else
'            sVencidos = ""
'        End If
'        VSQL = "Select C.cCodCta, C.dFecVenc, C.nValTasac, C.nTasaIntVenc, C.nCostCusto, C.cEstado, C.nSaldoCap, " _
'            & "C.nTasaInt, C.nPlazo, " _
'            & "nTasaCustodia = (Select nValor1 From Parametro Where cCodPar = '10102'), " _
'            & "nTasaImpuesto = (Select nValor1 From Parametro Where cCodPar = '10104'), " _
'            & "nCostoPrepRemate = (Select nValor1 From Parametro Where cCodPar = '10107') " _
'            & "From CredPrenda C Where C.cEstado IN ('1','4','6','7') " & sVencidos & " " _
'            & " And C.cCodCta IN " & sCuentas & " Order by C.cCodCta"
'
'        rsCta.Open VSQL, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
'        Set rsCta.ActiveConnection = Nothing
'        If Not (rsCta.EOF And rsCta.BOF) Then
'            Do While Not rsCta.EOF
'                nCuentasPrd = nCuentasPrd + 1
'                nDeudaPrend = CalculaDeudaPrendario(rsCta("nSaldoCap"), rsCta("dFecVenc"), rsCta("nValTasac"), rsCta("nTasaIntVenc"), _
'                                    rsCta("nCostCusto"), rsCta("nTasaImpuesto"), rsCta("cEstado"), rsCta("nCostoPrepRemate"))
'                nMinimoPagar = CalculaMinimoPagar(rsCta("nTasaInt"), rsCta("nPlazo"), rsCta("nValTasac"), rsCta("nTasaCustodia"), _
'                                    rsCta("nSaldoCap"), rsCta("cEstado"), rsCta("dFecVenc"), rsCta("nTasaIntVenc"), rsCta("nTasaCustodia"), _
'                                    rsCta("nTasaImpuesto"), rsCta("nCostoPrepRemate"))
'                rsCred.AddNew
'                rsCred("Cuenta") = Trim(rsCta("cCodCta"))
'                rsCred("Estado") = IIf(rsCta("cEstado") = "6", "RE", "")
'                rsCred("Vencimiento") = rsCta("dFecVenc")
'                rsCred("CuoVenc") = 1
'                rsCred("Cuota") = rsCta("nSaldoCap")
'                rsCred("Mora") = nDeudaPrend - rsCta("nSaldoCap")
'                rsCred("Gastos") = 0
'                rsCred("MinPagar") = nMinimoPagar
'                rsCred.Update
'                rsCta.MoveNext
'            Loop
'        End If
'        rsCta.Close
'    End If
'
'Next I
'
'nCuentasJud = 0
'
'If UBound(CuentasPersonaJud, 1) > 0 Then
'    sCuentas = "("
'    For I = 1 To UBound(CuentasPersonaJud, 1)
'        sCuentas = sCuentas & "'" & CuentasPersonaJud(I) & "',"
'    Next I
'    sCuentas = Left(sCuentas, Len(sCuentas) - 1)
'    sCuentas = sCuentas & ")"
'    If AbreConeccion(sAgencia & "0000000000") Then
'        'Verificamos los Creditos Judiciales concentrados en la agencia principal
'        VSQL = "Select C.cCodCta, C.cCondicion From CredCJudi C " _
'            & "Where C.cEstado = 'V' And C.cCodCta IN " & sCuentas
'        rsCta.Open VSQL, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
'        Set rsCta.ActiveConnection = Nothing
'        If Not (rsCta.EOF And rsCta.BOF) Then
'            Do While Not rsCta.EOF
'                nCuentasJud = nCuentasJud + 1
'                rsCred.AddNew
'                rsCred("Cuenta") = Trim(rsCta("cCodCta"))
'                rsCred("Estado") = "JU"
'                rsCred("Vencimiento") = Null
'                rsCred("CuoVenc") = 0
'                rsCred("Cuota") = 0
'                rsCred("Mora") = 0
'                rsCred("Gastos") = 0
'                rsCred("MinPagar") = 0
'                rsCred.Update
'                rsCta.MoveNext
'            Loop
'        End If
'        rsCta.Close
'    End If
'End If
'
'If (nCuentasCre + nCuentasPrd + nCuentasJud) = 0 Then
'    CreditosPendientes = False
'Else
'    CreditosPendientes = True
'End If

End Function


Public Function GeneraConvenioOrdenPago(ByVal sCuenta As String, ByVal dFecha As Date) As String
Dim rsPers As Recordset
Dim lsTit As String, lsRep As String
Dim lsDireccion As String, lsDNI As String
Dim sCad As String, sPersona As String, saux As String
Dim clsMant As DCapMantenimiento

Set clsMant = New DCapMantenimiento
Set rsPers = clsMant.GetPersonaCuenta(sCuenta)
Set clsMant = Nothing

Open App.path & "\FormatoCarta\ConOP.txt" For Input As #1
Do While Not EOF(1)
    Line Input #1, saux
    sCad = sCad & Chr$(10) & saux
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
                    lsRep = lsRep & Chr(10) & Space(40) & PstaNombre(Trim(rsPers("Nombre")))
                End If
            Else
                If lsTit = "" Then
                    lsDNI = rsPers("ID") & " " & rsPers("ID N°")
                    lsTit = PstaNombre(rsPers("Nombre"))
                    lsDireccion = rsPers("Direccion") & ""
                Else
                    lsTit = lsTit & Chr(10) & Space(52) & PstaNombre(rsPers("Nombre"))
                End If
            End If
            sPersona = rsPers("cPersCod")
        End If
        rsPers.MoveNext
    Loop
    
    'Datos de la Orden------------------------------------------------------------------------------------------------
      
    Dim lsInicio As String
    Dim lsFin As String
    Dim lsFecha As String
    Dim lsNumTal As String
    Dim lsPersNombre As String
    Dim lsTipo As String
    Dim lsMovNro As String
    
    Dim oMant As NCapMantenimiento
    Set oMant = New NCapMantenimiento
    
    Dim lrsOrdenPago As ADODB.Recordset
    Set lrsOrdenPago = New ADODB.Recordset
    
    
    Set lrsOrdenPago = oMant.GetOrdenPagoEmitidas(sCuenta)
                            
                            
                            
                            
    lsInicio = lrsOrdenPago("nInicio")
    lsFin = lrsOrdenPago("nFin")
    lsFecha = Format$(lrsOrdenPago("Fecha"), "dd/mm/yyyy")
    
'    lsNumTal = lrsOrdenPago("nNumTal")
'    lsPersNombre = PstaNombre(lrsOrdenPago("cPersNombre"), False)
'    lsTipo = lrsOrdenPago("nTipo")
'    lsMovNro = lrsOrdenPago("cMovNro")
    
    '-----------------------------------------------------------------------------------------------------------------
    sCad = Replace(sCad, "<<Inicio>>", Format(lsInicio, "#########"))
    sCad = Replace(sCad, "<<Fin>>", Format(lsFin, "#########"))
    sCad = Replace(sCad, "<<Fecha>>", lsFecha)
    sCad = Replace(sCad, "<<NumTal>>", Format(lsNumTal, "#########"))
    sCad = Replace(sCad, "<<PersNombre>>", lsPersNombre)
    sCad = Replace(sCad, "<<Tipo>>", Format(lsTipo, "#########"))
    sCad = Replace(sCad, "<<MovNro>>", lsMovNro)
    
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
GeneraConvenioOrdenPago = sCad
End Function

Private Sub GeneraConvenioOP(sCuenta As String)

    
    On Error GoTo GeneraMovNroErr
    
    Dim rs As Recordset
    Dim sql As String
    Dim oConn As DConecta
    
    Set oConn = New DConecta
    
    If oConn.AbreConexion = False Then Exit Sub
    sql = "SP_CapConvenioOrdenPago '" & sCuenta & "'"
    Set rs = oConn.Ejecutar(sql)
    Set rs = Nothing
    
    Exit Sub
GeneraMovNroErr:
    Call RaiseError(MyUnhandledError, "GeneraConvenioOP")
                
End Sub

Private Sub EliminaConvenioOP()

    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As DConecta
    
    Set rs = New ADODB.Recordset
    Set oConect = New DConecta
            
                    
    If oConect.AbreConexion = False Then Exit Sub
    
    sql = "Delete From tmpConvenioOP"
    
    Set rs = oConect.CargaRecordSet(sql)
    
    oConect.CierraConexion
    Set oConect = Nothing

End Sub

