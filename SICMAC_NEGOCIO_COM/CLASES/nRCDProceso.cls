VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nRCDProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Conecta As DConecta

Public Event ShowProgress()
Public Event CloseProgress()
Public Event Progress(pnValor As Long, pnTotal As Long)


Public Function nCargaParametroRCD(ByVal psFecha As String, ByVal psServConsol As String) As ADODB.Recordset
Dim lsSQL As String
Dim lrs As ADODB.Recordset
Dim loBase As DConecta
lsSQL = "Select * from " & psServConsol & "RCDParametro Where cMes ='" & psFecha & "' "
Set loBase = New DConecta
    loBase.AbreConexion
    Set lrs = loBase.CargaRecordSet(lsSQL)
Set loBase = Nothing
Set nCargaParametroRCD = lrs
End Function

' ** Crea las Tabla de Trabajo del RCD
Public Sub nCreaTablasRCD(ByVal psFecha As String, ByVal psServConsol As String)
Dim lsSQL As String
Dim loBase As DConecta
Set loBase = New DConecta
    loBase.AbreConexion

If nExisteTabla("RCDvc" & psFecha & "01", psServConsol) = False Then
    lsSQL = "CREATE TABLE " & psServConsol & "RCDvc" & psFecha & "01  ( " _
        & "cTipoFor char (1) NOT NULL ,cTipoInf char (1) NOT NULL ," _
        & "cNumSec char (6) NULL , cCodSBS char (10) NULL, " _
        & "cPersCod char (13) NOT NULL, " _
        & "cActEcon char (4) NULL , " _
        & "cCodRegPub char (15) NULL , " _
        & "cTidoTr char (1) NULL , " _
        & "cNudoTr char (11) NULL , " _
        & "cTiDoci char (1) NULL , " _
        & "cNuDoci char (12) NULL , " _
        & "cTipPers char (1) NULL , " _
        & "cResid char (1) NULL , " _
        & "cCalifica char (1) NULL , " _
        & "cMagEmp char (1) NULL , " _
        & "cAccionista char (1) NULL , " _
        & "cRelInst char (1) NULL , " _
        & "cPaisNac char (4) NULL, " _
        & "cSiglas char (20) NULL , " _
        & "cPersNom Varchar (120) NULL , " _
        & "cPersNomCom Varchar (120) NULL , " _
        & "CAPEPAT   VarChar (120) NULL,   " _
        & "CAPEMAT   VarChar (40) NULL, " _
        & "CAPECAS   VarChar (40) NULL,   " _
        & "CNOMBRE1  VarChar (40) NULL, " _
        & "CNOMBRE2  VarChar (40) NULL, "
lsSQL = lsSQL & "cPersGenero char (1) NULL , " _
        & "cPersEstado char (1) NULL , " _
        & "cIndRCC     char (1) NULL , " _
        & "CONSTRAINT PK_RCDvc" & psFecha & "01_4__10 PRIMARY KEY (cPersCod) )"
        
    loBase.Ejecutar (lsSQL)
     

    lsSQL = "CREATE  INDEX cNumSec ON " & psServConsol & "RCDvc" & psFecha & "01 (cNumSec)"
    loBase.Ejecutar (lsSQL)
    
    lsSQL = "CREATE  INDEX cNomPers ON " & psServConsol & "RCDvc" & psFecha & "01 (cPersNom)"
    loBase.Ejecutar (lsSQL)
    
    ' Saldos de Persona
    lsSQL = "CREATE TABLE " & psServConsol & "RCDvc" & psFecha & "02 ( " _
        & "cTipoFor char (1) NOT NULL , " _
        & "cTipoInf char (1) NOT NULL , " _
        & "cNumSec char (6) NULL , " _
        & "cCodAge char (4) NULL , " _
        & "cUbicGeo char (6) NULL , " _
        & "cCtaCnt char (14) NULL , " _
        & "cTipoCred char (1) NULL , " _
        & "nSaldo money NULL ,  " _
        & "nCondDias int NULL , " _
        & "cPersCod char (13) NOT NULL ) "
    
    loBase.Ejecutar (lsSQL)
    
    lsSQL = "CREATE INDEX [cCtaCnt_cTipoCred_nCondDias] ON " & psServConsol & "RCDvc" & psFecha & "02 ([cCtaCnt], [cTipoCred], [nCondDias])"
    loBase.Ejecutar (lsSQL)
    
    lsSQL = "CREATE INDEX [cNumSec] ON " & psServConsol & "RCDvc" & psFecha & "02 ([cCtaCnt])"
    loBase.Ejecutar (lsSQL)
    
    lsSQL = "CREATE INDEX [cCodPers_cCtaCont_nCondDias] ON " & psServConsol & "RCDvc" & psFecha & "02 ([cPersCod], [cCtaCnt], [nCondDias])"
    loBase.Ejecutar (lsSQL)
    
    ' *** Saldos de Cuentas de CMACT
    lsSQL = "CREATE TABLE " & psServConsol & "RCDvc" & psFecha & "03 (" _
        & "cTipoFor char (1) NOT NULL , " _
        & "cTipoInf char (1) NOT NULL , " _
        & "cNumSec char (6) NULL , " _
        & "cCtaCnt char (14) NULL , " _
        & "cTipoCred char (1) NULL , " _
        & "nSaldo money NULL ,  " _
        & "nCondDias int NULL  ) "
        
    loBase.Ejecutar (lsSQL)
    
    lsSQL = "CREATE INDEX [cCtaCnt_nCondDias] ON " & psServConsol & "RCDvc" & psFecha & "03 ([cCtaCnt], [nCondDias])"
    loBase.Ejecutar (lsSQL)
    
    lsSQL = "CREATE INDEX [cCtaCnt_cTipoCred_nCondDias] ON " & psServConsol & "RCDvc" & psFecha & "03 ([cCtaCnt], [cTipoCred], [nCondDias])"
    loBase.Ejecutar (lsSQL)
Else
    lsSQL = "DELETE FROM " & psServConsol & "RCDvc" & psFecha & "01"
    loBase.Ejecutar (lsSQL)

    lsSQL = "DELETE FROM " & psServConsol & "RCDvc" & psFecha & "02"
    loBase.Ejecutar (lsSQL)
    
    lsSQL = "DELETE FROM " & psServConsol & "RCDvc" & psFecha & "03"
    loBase.Ejecutar (lsSQL)
End If

Set loBase = Nothing
End Sub

'** Verifica si existen las tablas
Public Function nExisteTabla(ByVal lsNombreTabla As String, ByVal psServConsol As String) As Boolean
Dim rs As New ADODB.Recordset
Dim lsSQL As String
Dim loBase As DConecta

nExisteTabla = False
lsSQL = "Select * From " & psServConsol & "SysObjects Where Name = '" & lsNombreTabla & "' and type ='U'"
Set loBase = New DConecta
    loBase.AbreConexion
    Set rs = loBase.CargaRecordSet(lsSQL)
Set loBase = Nothing
If Not (rs.BOF And rs.EOF) Then
    nExisteTabla = True
End If
rs.Close
Set rs = Nothing
End Function

Public Function nObtieneCalificacionPersonaProcesada(ByVal psPersona As String, psServConsol As String) As String
Dim lsSQL As String
Dim rs As ADODB.Recordset
Dim loBase As DConecta
Dim lsCalGen As String
lsCalGen = "0"
lsSQL = "Select cCalGen FROM ColocCalifProv WHERE cPersCod ='" & Trim(psPersona) & "'"
Set loBase = New DConecta
    loBase.AbreConexion
    Set rs = loBase.CargaRecordSet(lsSQL)
    
    If rs.BOF And rs.EOF Then
        MsgBox "No se encontro Calificacion de Persona", vbInformation, "Aviso"
    Else
        If IsNull(rs!cCalGen) Then
            MsgBox "No se ha asignado calificacion Correcta", vbInformation, "Aviso"
        Else
            lsCalGen = Trim(rs!cCalGen)
        End If
    End If
    rs.Close
    Set rs = Nothing
Set loBase = Nothing
nObtieneCalificacionPersonaProcesada = lsCalGen
End Function


Public Function FormatoCtaContable(ByVal pCuentaCnt As String) As String
  FormatoCtaContable = Trim(pCuentaCnt) & String(14 - Len(Trim(pCuentaCnt)), "0")
End Function
Public Function EmiteCodigoPersona(lsCodAux As String) As String
Dim Co As DConecta
Dim sql As String
Dim rs As New ADODB.Recordset
Set Co = New DConecta
sql = "Select * from RCDCodigoAux where cCodAux='" & Trim(lsCodAux) & "'"
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
'Rs.Open SQl, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not RSVacio(rs) Then
    EmiteCodigoPersona = Trim(rs!cCodPers)
Else
    EmiteCodigoPersona = ""
End If
rs.Close
Set rs = Nothing
Set Co = Nothing
End Function


Private Function EmiteCalificacion(lsCodPers As String) As String
Dim Co As DConecta
Dim sql As String
Dim rs As New ADODB.Recordset
Set Co = New DConecta
sql = "Select * from ColocevalCalif where cperscod='" & lsCodPers & "'"
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
EmiteCalificacion = ""
If Not RSVacio(rs) Then
    EmiteCalificacion = Trim(rs!cEvalCalif)
End If
rs.Close
Set rs = Nothing
End Function

Private Sub Class_Initialize()
Set Conecta = New DConecta
End Sub


Private Sub Class_Terminate()
Set Conecta = Nothing
End Sub




