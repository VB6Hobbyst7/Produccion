VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMRCDReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Conecta As comconecta.DCOMConecta

Function GetQuery(ByVal Sql As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
    Conecta.AbreConexion
    Set rs = Conecta.CargaRecordSet(Sql)
        Conecta.CierraConexion
    Set GetQuery = rs
End Function

Public Function GetServerConsol() As String
Dim rs As New ADODB.Recordset
Set rs = GetQuery("select nConsSisValor from ConstSistema where nConsSisCod=43")
GetServerConsol = IIf(IsNull(rs!nConsSisValor), "", rs!nConsSisValor)
Set rs = Nothing
End Function

Public Function ExisteTabla(ByVal psServerConsol As String, ByVal psTabla As String) As ADODB.Recordset
Dim Sql As String
Dim rs As New ADODB.Recordset
Sql = "select * from " & psServerConsol & "sysobjects where name = '" & psTabla & "'"
Set rs = GetQuery(Sql)
Set ExisteTabla = rs
Set rs = Nothing
End Function

Public Function nRepo179201_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "   Select cPersCod,cPersNom,cCalifica as Total " _
            & " from " & psServerConsol & "RCDvc" & psFecha & "01 where (cCalifica ='' or cCalifica is null or cCalifica not in('0','1','2','3','4'))"

Set rs = GetQuery(Sql)
Set nRepo179201_ = rs
Set rs = Nothing
End Function

Public Function nRepo179202_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom, cTipPers, cCodSBS from " & psServerConsol & "RCDvc" & psFecha & "01  " _
        & " WHERE (PATINDEX('%Y/O%', cPersNom )<>0  OR  PATINDEX('% O %', cPersNom )<>0 OR PATINDEX('% Y %', cPersNom)<>0 )  and cTipPers<>'3' " _
        & " ORDER BY cTipPers"
Set rs = GetQuery(Sql)
Set nRepo179202_ = rs
Set rs = Nothing
End Function

Public Function nRepo179203_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cCodSBS, cPersNom from " & psServerConsol & "RCDvc" & psFecha & "01  " _
        & " WHERE (cCodSbs IS NULL OR cCodSbs ='' or LEN(cCodSbs)<>10) "
Set rs = GetQuery(Sql)
Set nRepo179203_ = rs
Set rs = Nothing
End Function

Public Function nRepo179204_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cPersNom,cTipPers,cMagEmp,cNudoTr FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE CTIPPERS ='2' AND CMAGEMP  NOT IN ('1','2','3','4') "
Set rs = GetQuery(Sql)
Set nRepo179204_ = rs
Set rs = Nothing
End Function

Public Function nRepo179205_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cPersNom,cTipPers,cMagEmp,cNudoci FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE CTIPPERS NOT IN ('2') AND CMAGEMP IN ('1','2','3','4')"
Set rs = GetQuery(Sql)
Set nRepo179205_ = rs
Set rs = Nothing
End Function

Public Function nRepo179206_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cPersNom,cTipPers,cMagEmp , cNudoci, cNudotr  FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (CMAGEMP IS NULL  OR CMAGEMP ='' OR CMAGEMP NOT IN ('0','1','2','3','4'))"
Set rs = GetQuery(Sql)
Set nRepo179206_ = rs
Set rs = Nothing
End Function

Public Function nRepo179207_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cTidoci, cNudoci, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01  WHERE cTidoci not in ('1','2','3','4','5') and cTipPers in ('1','3')"
Set rs = GetQuery(Sql)
Set nRepo179207_ = rs
Set rs = Nothing
End Function

Public Function nRepo179208_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cPersNom,cTipPers, cNudoci, cCodSbS FROM " & psServerConsol & "RCDvc" & psFecha & "01  WHERE (cNudoCi like '%000000%' or cNudoci = '' or cNudoci is null)  and cTipPers in ('1','3')"
Set rs = GetQuery(Sql)
Set nRepo179208_ = rs
Set rs = Nothing
End Function

Public Function nRepo179209_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
 Sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cNudoTr,cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (cNudoTr like '%000000%' or cNudoTr = '' or cNudoTr is null)  and cTipPers in ('2')"
Set rs = GetQuery(Sql)
Set nRepo179209_ = rs
Set rs = Nothing
End Function

Public Function nRepo179210_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cTidotr, cNudotr, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE cTidotr not in ('2','3') and cTipPers in ('2')"
Set rs = GetQuery(Sql)
Set nRepo179210_ = rs
Set rs = Nothing
End Function

Public Function nRepo179211_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (cTipPers not in ('1','2','3','4') or cTipPers is null or cTipPers='')"
Set rs = GetQuery(Sql)
Set nRepo179211_ = rs
Set rs = Nothing
End Function

Public Function nRepo179212_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cCodSbs,cActEcon  FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE len(cActeCon)<>4 or cActecon like '%000%' or cActecon ='' or cActecon is Null"
Set rs = GetQuery(Sql)
Set nRepo179212_ = rs
Set rs = Nothing
End Function

Public Function nRepo179213_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cNudoTr, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (cSiglas is Not Null or cSiglas<>'') "
Set rs = GetQuery(Sql)
Set nRepo179213_ = rs
Set rs = Nothing
End Function

Public Function nRepo179214_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cCodRegPub, cNudoTr, cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (cCodRegPub like '%00000000000%' or cCodRegPub is null or cCodRegPub ='' ) and cTipPers ='2'"
Set rs = GetQuery(Sql)
Set nRepo179214_ = rs
Set rs = Nothing
End Function

Public Function nRepo179215_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod, cPersNom   FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE  CTIPPERS IN ('1') ORDER BY cPersCod "
Set rs = GetQuery(Sql)
Set nRepo179215_ = rs
Set rs = Nothing
End Function

Public Function nRepo179216_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cNudoTr,cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (len(cNudoTr)<>11 and cTidoTr='2' )  and cTipPers in ('2')  "
Set rs = GetQuery(Sql)
Set nRepo179216_ = rs
Set rs = Nothing
End Function

Public Function nRepo179217_(ByVal psServerConsol As String, ByVal psFecha As String) As ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT cPersCod,cPersNom,cTipPers, cSiglas, cNudoTr,cCodSbs FROM " & psServerConsol & "RCDvc" & psFecha & "01 WHERE (len(cNudoCi)<>8 and cTidoci='1' )  and cTipPers in ('1','3')  "
Set rs = GetQuery(Sql)
Set nRepo179217_ = rs
Set rs = Nothing
End Function

Public Function nRepo179101_ArchivoRCD(ByVal psServerConsol As String, ByVal pdFecSis As Date, ByRef psMensaje As String) As String
Dim SQL1 As String
Dim Sql2 As String
Dim CadPrint() As String
Dim lsCadenaAux As String
Dim lsCadenaPrint As String
Dim TotalCadenas As Long
Dim contTotal As Long
Dim Cont As Long

Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset

Dim lsApellPat As String
Dim lsApellMat As String
Dim lsPrimerNombre As String
Dim lsSegundoNombre As String
Dim lsApellCony As String
Dim lsNombrePers As String
Dim Pos1 As Long
Dim Pos2 As Long
Dim Pos3 As Long
Dim lsCadAux As String
Dim lscadena As String
Dim lsTipPer As String
Dim lsGenero As String
Dim lsEstado As String
Dim lsCondCta As String
Dim lsCondDisp As String

'''
Dim J As Long
Dim lscTidoci As String
Dim lscNudoCi As String
Dim lscTidoTr As String
Dim lscNudotr As String
Dim lsActividad As String
Dim lsCodSbs As String
Dim lsCalificacion As String
Dim lsCodPers As String

Dim lsPersonaAct As String
Dim lnSaldo As Currency

Dim lsNumSecICC As String
Dim Conecta As comconecta.DCOMConecta
Set Conecta = New comconecta.DCOMConecta
Dim lsIndRCC As String
' Crea el nombre de archivo.

Dim lsCodOfiInfICC As String
Dim lsUbicaGeoICC As String
Dim lsCodDptoICC As String

Dim gsCodEmpInfICC As String
Dim gnMontoMinimoICC As Integer

Dim ofuni As New COMFunciones.FCOMImpresion
Dim ofunc As New COMFunciones.FCOMCadenas
Dim lsCadImp As String

SQL1 = "select * from constsistema where nConsSisCod=80"
Set rs1 = GetQuery(SQL1)
gsCodEmpInfICC = rs1!nConsSisValor

SQL1 = "select top 1 * from " & psServerConsol & "rcdparametro  order by  cMEs desc"
Set rs1 = GetQuery(SQL1)
gnMontoMinimoICC = rs1!nMontoMin
'--COPIAR--
'lsNombreArchivo = App.Path & "\Spooler\RCD" & Format(pdFecSis, "yyyymm") & ".108"
'Open lsNombreArchivo For Output As #NúmeroArchivo

'**** Registro CABECERA
lsCadImp = lsCadImp & "0106" & "01" & ofunc.FillNum(Trim(gsCodEmpInfICC), 5, "0") & Format(pdFecSis, "yyyymmdd") & "012" _
        & Space(15) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(gnMontoMinimoICC), 0, gnMontoMinimoICC)), 15, 0, False) & Chr(13) & Chr(10)

SQL1 = "SELECT * FROM " & psServerConsol & "sysobjects WHERE name = '" & "RCDvc" & Format(pdFecSis, "yyyymm") & "01'"
Conecta.AbreConexion
Set rs1 = Conecta.CargaRecordSet(SQL1)

If rs1.EOF And rs1.BOF Then
    Conecta.CierraConexion
    psMensaje = "Tabla [RCDvc" & Format(pdFecSis, "yyyymm") & "01]  No existe, Por favor comuniquese con el Dpto de Sistemas"
    Exit Function
End If
      
SQL1 = " Select COUNT(*) AS TOTAL From " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "01 R1 " _
     & " Inner Join " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "02 R2 ON R2.cPersCod=R1.cPersCod "

Set rs1 = Conecta.CargaRecordSet(SQL1)
Conecta.CierraConexion

If Not (rs1.EOF And rs1.BOF) Then
    contTotal = IIf(IsNull(rs1!Total), 0, rs1!Total)
End If
rs1.Close
Set rs1 = Nothing

SQL1 = "Select R1.cTipoFor, R1.cTipoInf, R1.cNumSec, R1.cCodSBS, R1.cPersCod, R1.cActEcon, " _
    & "   R1.cCodRegPub, R1.cTidoTr, R1.cNudoTr, R1.cTiDoci, R1.cNuDoci, " _
    & "   R1.cTipPers, R1.cResid, R1.cCalifica, R1.cMagEmp, R1.cAccionista, R1.cRelInst,  " _
    & "   R1.cPaisNac, R1.cSiglas, R1.cPersNom,R1.cPersNomCom,R1.cPersGenero,R1.cPersEstado, R2.cTipoFor cTipoFor2, R2.cTipoInf cTipoInf2 , " _
    & "   R2.cCodAge, R2.cUbicGeo, R2.cCtaCnt, R2.cTipoCred, R2.nSaldo, R2.nCondDias,   " _
    & "   R2.cPersCod cCodPers2, CAPEPAT, CAPEMAT,CAPECAS,CNOMBRE1,CNOMBRE2, R1.cIndRCC " _
    & " From " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "01 R1 " _
    & "      Inner Join " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "02 R2 ON R2.cPersCod=R1.cPersCod " _
    & " ORDER BY CAPEPAT, CAPEMAT,CAPECAS,CNOMBRE1,CNOMBRE2 "

Conecta.AbreConexion
Set rs1 = Conecta.CargaRecordSet(SQL1)
Conecta.CierraConexion
Cont = 0: J = 0
'Screen.MousePointer = 11

If Not (rs1.EOF And rs1.BOF) Then
    
    Do While Not rs1.EOF
        J = J + 1
            
        If lsPersonaAct <> rs1!cPersCod Then

            Cont = Cont + 1
            lsNumSecICC = ofunc.FillNum(Trim(Str(Cont)), 6, "0")
            
            'lsNombrePers = oFunC.PSTANOMBRE(CadDerecha(Trim(Replace(rs1!cPersNom, "-", "", , , vbTextCompare)), 80))
            'lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "/", "", , , vbTextCompare)), 80)
            'lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ",", "", , , vbTextCompare)), 80)
            'lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ".", " ", , , vbTextCompare)), 80)
            'lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "Ñ", "#", , , vbTextCompare)), 80)
            'lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "ñ", "#", , , vbTextCompare)), 80)
            '''''''''''''''
            ''''''INSERCION DEL 11/06/2004 PARA SEPARAR NOMBRES  - ENME

            lsTipPer = rs1!cTipPers
            lsApellPat = "": lsApellMat = "": lsPrimerNombre = "": lsApellCony = "": lsSegundoNombre = ""
            lsNombrePers = rs1!cPersNomCom
            If Len(Trim(lsNombrePers)) = 0 Then
                    lsNombrePers = Trim(lsNombrePers)
            Else
                If InStr(1, lsNombrePers, "Y/O", vbTextCompare) = 0 And InStr(1, lsNombrePers, " Y ", vbTextCompare) = 0 And InStr(1, lsNombrePers, " O ", vbTextCompare) = 0 Then
                   lsNombrePers = Trim(lsNombrePers)
                End If
            End If
            lsNombrePers = Trim(Replace(lsNombrePers, "-/", "", , , vbTextCompare))
            lsNombrePers = Trim(Replace(lsNombrePers, "-", "", , , vbTextCompare))
            lsNombrePers = Trim(Replace(lsNombrePers, ".", " ", , , vbTextCompare))
            lsNombrePers = Trim(Replace(lsNombrePers, "//", "/", , , vbTextCompare))
            
            '-****Separa Nombre de Clientes
'                If InStr(1, lsNombrePers, "Y/O", vbTextCompare) = 0 And InStr(1, lsNombrePers, " Y ", vbTextCompare) = 0 And InStr(1, lsNombrePers, " O ", vbTextCompare) = 0 Then
'                    If Trim(lsTipPer) = "1" Then
'                        Pos1 = InStr(1, lsNombrePers, "/", vbTextCompare)
'                        Pos2 = InStr(1, lsNombrePers, "\", vbTextCompare)
'                        Pos3 = InStr(1, lsNombrePers, ",", vbTextCompare)
'                        If Pos1 > 0 Then
'                            lsApellPat = Mid(lsNombrePers, 1, Pos1 - 1)
'                            lsCadAux = Mid(lsNombrePers, Pos1 + 1, Len(lsNombrePers))
'                            Pos1 = InStr(1, lsCadAux, "\", vbTextCompare)
'                            If Pos1 > 0 Then
'                                lsApellMat = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                lsCadAux = Mid(lsCadAux, Pos1 + 1, Len(lsCadAux))
'                                If Len(Trim(lsApellMat)) = 0 Then
'                                    Pos1 = InStr(1, lsCadAux, ",", vbTextCompare)
'                                    If Pos1 > 0 Then
'                                        lsApellCony = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                        lsCadAux = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                        If Len(lsCadAux) > 0 Then
'                                            Pos1 = InStr(1, lsCadAux, " ", vbTextCompare)
'                                            If Pos1 > 0 Then
'                                                lsPrimerNombre = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                                lsCadAux = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                                If Len(lsCadAux) > 0 Then
'                                                    lsSegundoNombre = lsCadAux
'                                                End If
'                                            Else
'                                                lsPrimerNombre = lsCadAux
'                                            End If
'
'                                        End If
'
'                                    Else
'                                        lsApellCony = lsCadAux
'                                    End If
'                                Else
'                                    Pos1 = InStr(1, lsCadAux, ",", vbTextCompare)
'                                    If Pos1 > 0 Then
'                                        lsApellCony = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                        lsCadAux = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                        If Len(lsCadAux) > 0 Then
'                                            Pos1 = InStr(1, lsCadAux, " ", vbTextCompare)
'                                            If Pos1 > 0 Then
'                                                lsPrimerNombre = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                                lsCadAux = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                                If Len(lsCadAux) > 0 Then
'                                                    lsSegundoNombre = lsCadAux
'                                                End If
'                                            Else
'                                                lsPrimerNombre = lsCadAux
'                                            End If
'
'                                        End If
'
'                                    Else
'                                        lsApellCony = lsCadAux
'                                    End If
'
'                                End If
'                            Else
'                                Pos1 = InStr(1, lsCadAux, ",", vbTextCompare)
'                                If Pos1 > 0 Then
'                                    lsApellMat = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                    lsCadAux = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                        If Len(lsCadAux) > 0 Then
'                                            Pos1 = InStr(1, lsCadAux, " ", vbTextCompare)
'                                            If Pos1 > 0 Then
'                                                lsPrimerNombre = Trim(Mid(lsCadAux, 1, Pos1 - 1))
'                                                lsCadAux = Trim(Mid(lsCadAux, Pos1 + 1, Len(lsCadAux)))
'                                                If Len(lsCadAux) > 0 Then
'                                                    lsSegundoNombre = lsCadAux
'                                                End If
'                                            Else
'                                                lsPrimerNombre = lsCadAux
'                                            End If
'                                        End If
'                                End If
'                            End If
'                       ElseIf Pos2 > 0 Then
'                            lsApellPat = Mid(lsNombrePers, 1, Pos2 - 1)
'                            lsCadAux = Mid(lsNombrePers, Pos2 + 1, Len(lsNombrePers))
'                            lsApellMat = ""
'                                If Len(Trim(lsApellMat)) = 0 Then
'                                    Pos2 = InStr(1, lsCadAux, ",", vbTextCompare)
'                                    If Pos2 > 0 Then
'                                        lsApellCony = Trim(Mid(lsCadAux, 1, Pos2 - 1))
'                                        lsCadAux = Trim(Mid(lsCadAux, Pos2 + 1, Len(lsCadAux)))
'                                        If Len(lsCadAux) > 0 Then
'                                            Pos2 = InStr(1, lsCadAux, " ", vbTextCompare)
'                                            If Pos2 > 0 Then
'                                                lsPrimerNombre = Trim(Mid(lsCadAux, 1, Pos2 - 1))
'                                                lsCadAux = Trim(Mid(lsCadAux, Pos2 + 1, Len(lsCadAux)))
'                                                If Len(lsCadAux) > 0 Then
'                                                    lsSegundoNombre = lsCadAux
'                                                End If
'                                            Else
'                                                lsPrimerNombre = lsCadAux
'                                            End If
'
'                                        End If
'
'                                    Else
'                                        lsApellCony = lsCadAux
'                                    End If
'                                End If
'                        ElseIf Pos3 > 0 Then
'                            lsApellPat = Mid(lsNombrePers, 1, Pos3 - 1)
'                            lsCadAux = Mid(lsNombrePers, Pos3 + 1, Len(lsNombrePers))
'                            If Len(lsCadAux) > 0 Then
'                                 Pos3 = InStr(1, lsCadAux, " ", vbTextCompare)
'                                 If Pos3 > 0 Then
'                                     lsPrimerNombre = Trim(Mid(lsCadAux, 1, Pos3 - 1))
'                                     lsCadAux = Trim(Mid(lsCadAux, Pos3 + 1, Len(lsCadAux)))
'                                     If Len(lsCadAux) > 0 Then
'                                         lsSegundoNombre = lsCadAux
'                                     End If
'                                 Else
'                                     lsPrimerNombre = lsCadAux
'                                 End If
'
'                            End If
'                        End If
'                        ''''''''''''''''
'                        If Pos1 = 0 And Pos2 = 0 And Pos3 = 0 Then
'                            lsApellPat = lsNombrePers
'                            lsApellMat = ""
'                            lsPrimerNombre = ""
'                        End If
'                    Else
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "-", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "$", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "%", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "/", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ")", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "(", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "@", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "*", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "+", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "  ", " ", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "   ", " ", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "\", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ",", "", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ".", " ", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "Ñ", "#", , , vbTextCompare)), 80)
'                        lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "ñ", "#", , , vbTextCompare)), 80)
'                        lsApellPat = lsNombrePers
'                        lsApellMat = ""
'                        lsPrimerNombre = ""
'                    End If
'                Else
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "-", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "$", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "%", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "/", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ")", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "(", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "@", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "*", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "+", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "  ", " ", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "   ", " ", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "\", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ",", "", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, ".", " ", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "Ñ", "#", , , vbTextCompare)), 80)
'                    lsNombrePers = CadDerecha(Trim(Replace(lsNombrePers, "ñ", "#", , , vbTextCompare)), 80)
'                    lsApellPat = lsNombrePers
'                    lsApellMat = ""
'                    lsPrimerNombre = ""
'                    lsSegundoNombre = ""
'                End If
                
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Not IsNull(rs1!cPersGenero) Then
                lsGenero = Trim(rs1!cPersGenero)
            Else
                lsGenero = ""
            End If
            If Not IsNull(rs1!cPersEstado) Then
                Select Case Trim(rs1!cPersEstado)
                    Case "2", "5", "6"
                        lsEstado = "C"
                    Case "3"
                        lsEstado = "V"
                    Case "4"
                         lsEstado = "D"
                    Case "1"
                         lsEstado = "S"
                    Case Else
                         lsEstado = Trim(rs1!cPersEstado)
                End Select
            Else
                lsEstado = ""
            End If
                    
            ''''''''''''FIN DE INSERCION DEL 11/06/2004
                        
            Select Case Trim(rs1!cTipPers)
                Case "1", "3"
                    lscTidoci = IIf(IsNull(rs1!ctidoci), "", IIf(Len(Trim(rs1!ctidoci)) = 0, "", Trim(rs1!ctidoci)))
                    lscNudoCi = IIf(IsNull(rs1!cNudoci), "", IIf(Len(Trim(rs1!cNudoci)) = 0, "", Trim(rs1!cNudoci)))
                    lscTidoTr = ""
                    lscNudotr = ""
                Case Else
                    lscTidoci = ""
                    lscNudoCi = ""
                    lscTidoTr = IIf(IsNull(rs1!cTidoTr), "", IIf(Len(Trim(rs1!cTidoTr)) = 0, "", Trim(rs1!cTidoTr)))
                    lscNudotr = IIf(IsNull(rs1!cNudoTr), "", IIf(Len(Trim(rs1!cNudoTr)) = 0, "", Trim(rs1!cNudoTr)))
                    ' RUC DE 11 DIGITOS
                    lscTidoTr = IIf(lscTidoTr = "2", "3", lscTidoTr)
            End Select
            If Not IsNull(rs1!cCodSBS) Then
                lsCodSbs = Trim(rs1!cCodSBS)
            Else
                lsCodSbs = ""
            End If
            If Not IsNull(rs1!cActEcon) Then
                If Len(Trim(rs1!cActEcon)) > 0 Then
                    lsActividad = ofunc.FillNum(rs1!cActEcon, 4, "0")
                Else
                    lsActividad = ""
                End If
            Else
                lsActividad = ""
            End If
            
            'lsCodPers = EmiteCodigoPersona(rs1!cPersCod, psServerConsol)
            'If Len(Trim(lsCodPers)) = 0 Then
                lsCodPers = Trim(rs1!cPersCod)
            'End If
            
            If lsCodPers = "1080800002739" Or lsCodPers = "1080100791482" Or lsCodPers = "1080100791482" Then
            '    Stop
            End If
            
            lsApellPat = Trim(rs1!CAPEPAT)
            lsApellMat = Trim(rs1!CAPEMAT)
            lsApellCony = Trim(rs1!CAPECAS)
            lsPrimerNombre = Trim(rs1!cNombre1)
            lsSegundoNombre = Trim(rs1!cNombre2)
            lsIndRCC = Trim(rs1!cIndRCC)
            
         lsCadImp = lsCadImp & rs1!cTipoFor & rs1!cTipoInf & ofunc.FillNum(Trim(lsNumSecICC), 8, " ") & _
                    ofuni.ImpreFormat(IIf(Len(Trim(lsCodSbs)) = 0, "", lsCodSbs), 10, 0) & ofuni.ImpreFormat(Trim(rs1!cPersCod), 20, 0) & ofunc.FillNum(lsActividad, 4, " ") & _
                    ofunc.FillNum(Trim(IIf(IsNull(rs1!ccodregpub), "", rs1!ccodregpub)), 15, " ") & _
                    ofunc.FillNum(IIf(Len(Trim(lscTidoTr)) = 0, "", lscTidoTr), 1, " ") & _
                    ofunc.FillNum(IIf(Len(Trim(lscNudotr)) = 0, "", lscNudotr), 11, " ") & ofunc.FillNum(IIf(Len(Trim(lscTidoci)) = 0, "", lscTidoci), 1, " ") & _
                    ofuni.ImpreFormat(IIf(Len(Trim(lscNudoCi)) = 0, "", lscNudoCi), 12, 0) & _
                    ofunc.FillNum(Trim(IIf(IsNull(rs1!cTipPers), "", rs1!cTipPers)), 1, " ") & ofunc.FillNum(Trim(IIf(IsNull(rs1!cResid), "", rs1!cResid)), 1, " ") & _
                    ofunc.FillNum(Trim(rs1!cCalifica), 1, " ") & _
                    ofunc.FillNum(Trim(IIf(IsNull(rs1!cMagEmp), "", rs1!cMagEmp)), 1, " ") & _
                    ofunc.FillNum(Trim(IIf(IsNull(rs1!cAccionista), "", rs1!cAccionista)), 1, " ") & ofunc.FillNum(Trim(IIf(IsNull(rs1!cRelInst), "", rs1!cRelInst)), 1, " ") & _
                    ofunc.FillNum(Trim(rs1!cPaisNac), 4, " ") & ofunc.FillNum(Trim(lsGenero), 1, " ") & ofunc.FillNum(Trim(lsEstado), 1, " ") & _
                    ofuni.ImpreFormat(Trim(IIf(IsNull(rs1!cSiglas), "", rs1!cSiglas)), 20, 0) & _
                    ofuni.ImpreFormat(Trim(lsApellPat), 120, 0) & _
                    ofuni.ImpreFormat(Trim(lsApellMat), 40, 0) & _
                    ofuni.ImpreFormat(Trim(lsApellCony), 40, 0) & _
                    ofuni.ImpreFormat(Trim(lsPrimerNombre), 40, 0) & _
                    ofuni.ImpreFormat(Trim(lsSegundoNombre), 40, 0) & ofuni.ImpreFormat(Trim(lsIndRCC), 1, 0) & Chr(13) & Chr(10)
         
        End If
        
        lsCadImp = lsCadImp & rs1!cTipoFor2 & rs1!cTipoInf2 & ofunc.FillNum(Trim(lsNumSecICC), 8, " ") & _
                ofunc.FillNum(Trim(rs1!CCodage), 4, "0") & ofunc.FillNum(Trim(rs1!cUbicGeo), 6, " ") & _
                FormatoCtaContable(rs1!cctacnt) & ofunc.FillNum(Trim(rs1!cTipoCred), 1, " ") & _
                ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs1!nSaldo), 0, rs1!nSaldo)), 15, 0, False) & _
                ofunc.FillNum(Trim(rs1!nCondDias), 4, " ") & "06" & "01" & Chr(13) & Chr(10)
        
        '********************* ACTUALIZANDO NUMERO DE SECUENCIA PARA REPORTE ICC *******************
        SQL1 = "Update " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "01 SET CNUMSEC='" & Trim(lsNumSecICC) & "' WHERE cPersCod='" & Trim(rs1!cPersCod) & "'"
        Sql2 = "Update " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "02 SET CNUMSEC='" & Trim(lsNumSecICC) & "' WHERE cPersCod='" & Trim(rs1!cPersCod) & "'"
        'Conecta.AbreConexion
        'Conecta.Ejecutar (SQL1)
        'Conecta.Ejecutar (Sql2)
        'Conecta.CierraConexion
        '*******************************************************************************************
        lsPersonaAct = rs1!cPersCod
        rs1.MoveNext
    Loop
    
End If
rs1.Close
Set rs1 = Nothing
Conecta.CierraConexion
'***************************************
'** Totales de la Empresa
'***************************************

Cont = Cont + 1
lsNumSecICC = ofunc.FillNum(Trim(Str(Cont)), 6, "0")

lsCadImp = lsCadImp & "2" & "1" & ofunc.FillNum(Trim(lsNumSecICC), 6, " ") & _
        Space(10) & Chr(13) & Chr(10)
        
SQL1 = "Select * from " & psServerConsol & "RCDvc" & Format(pdFecSis, "yyyymm") & "03 Order by cCtaCnt,cTipoCred,nCondDias "

Conecta.AbreConexion
Set rs1 = Conecta.CargaRecordSet(SQL1)
Conecta.CierraConexion
If Not (rs1.EOF And rs1.BOF) Then
    Do While Not rs1.EOF
        
        lsCadImp = lsCadImp & "2" & "2" & ofunc.FillNum(Trim(lsNumSecICC), 6, " ") & _
                ofunc.FillNum(Trim(lsCodOfiInfICC), 4, " ") & ofunc.FillNum(Trim(lsUbicaGeoICC), 6, " ") & FormatoCtaContable(rs1!cctacnt) & ofunc.FillNum(Trim(rs1!cTipoCred), 1, " ") & _
                ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs1!nSaldo), 0, rs1!nSaldo)), 15, 0, False) & _
                ofunc.FillNum(Trim(rs1!nCondDias), 4, " ") & Chr(13) & Chr(10)
        
        rs1.MoveNext
    Loop
End If
rs1.Close
Set rs1 = Nothing
Set ofunc = Nothing
Set ofuni = Nothing
End Function

Public Function FormatoCtaContable(ByVal pCuentaCnt As String) As String
  FormatoCtaContable = Trim(pCuentaCnt) & String(14 - Len(Trim(pCuentaCnt)), "0")
End Function

Public Function EmiteCodigoPersona(ByVal lsCodAux As String, ByVal psServerConsol As String) As String
Dim Co As comconecta.DCOMConecta
Dim Sql As String
Dim rs As New ADODB.Recordset
Set Co = New comconecta.DCOMConecta
Sql = "Select * from " & psServerConsol & "RCDCodigoAux where cCodAux='" & Trim(lsCodAux) & "'"
Co.AbreConexion
Set rs = Co.CargaRecordSet(Sql)
Co.CierraConexion

If Not (rs.EOF And rs.BOF) Then
    EmiteCodigoPersona = Trim(rs!cPersCod)
Else
    EmiteCodigoPersona = ""
End If
rs.Close
Set rs = Nothing
Set Co = Nothing
End Function

Private Function EliminaPunto(lnNumero As Currency) As Currency
Dim pos As Long
Dim CadAux As String
Dim CadAux1 As String
Dim lsNumero As String
lsNumero = Trim(Str(lnNumero))
If Val(lsNumero) > 0 Then
    pos = InStr(1, lsNumero, ".", vbTextCompare)
    If pos > 0 Then
        CadAux = Mid(lsNumero, 1, pos - 1)
        CadAux1 = Mid(lsNumero, pos + 1, Len(Trim(lsNumero)))
        If Len(Trim(CadAux1)) = 1 Then
            CadAux1 = CadAux1 & "0"
        End If
        EliminaPunto = CCur(CadAux & CadAux1)
    Else
        EliminaPunto = lnNumero & "00"
    End If
Else
    EliminaPunto = lnNumero
End If
End Function

Private Function EmiteCalificacion(lsCodPers As String) As String
Dim Co As comconecta.DCOMConecta
Dim Sql As String
Dim rs As New ADODB.Recordset
Set Co = New comconecta.DCOMConecta
Sql = "Select * from ColocevalCalif where cperscod='" & lsCodPers & "'"
Co.AbreConexion
Set rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
EmiteCalificacion = ""
If Not (rs.EOF And rs.BOF) Then
    EmiteCalificacion = Trim(rs!cEvalCalif)
End If
rs.Close
Set rs = Nothing
End Function

Public Sub nRepo179103_IBM(ByVal psServerConsol As String, ByVal dFecSis As Date, ByVal psNomCmact As String, ByVal psCodCmac As String)
Dim lsNombrePers As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim contTotal  As Long
Dim Sql As String
Dim J As Long
Dim lsApellPat As String
Dim lsApellMat As String
Dim lsApellCony As String
Dim lsNombre As String
Dim pos As Long
Dim Pos1 As Long
Dim Pos2 As Long
Dim Cont As Long
Dim lsTipoDoc As String
Dim lsNumDoc As String
Dim lsPrefijos(6) As String
Dim i As Long
Dim m As Long
Dim lsCodPers As String
Dim lsCadAux As String

Dim TOTALTIPODOC1 As Long: Dim TOTALTIPODOC2 As Long: Dim TOTALTIPODOC3 As Long
Dim TOTALTIPODOC4 As Long: Dim TOTALTIPODOC5 As Long: Dim TOTALTIPODOC6 As Long

Dim TOTALPERSNAT As Long: Dim TOTALPERSJUR As Long

Dim TOTALVENCMEN30MN As Long: Dim TOTALVENCMAY30MN As Long
Dim TOTALVENCMEN30ME As Long: Dim TOTALVENCMAY30ME As Long

Dim TOTALCAL0 As Long: Dim TOTALCAL1 As Long: Dim TOTALCAL2 As Long:
Dim TOTALCAL3 As Long: Dim TOTALCAL4 As Long

Dim TOTALNORMMN As Currency: Dim TOTALREFMN As Currency: Dim TOTALMEN30MN As Currency
Dim TOTALMAY30MN  As Currency: Dim TOTALCOBJUDMN  As Currency: Dim TOTALLINCREDMN As Currency:
Dim TOTALCASTMN As Currency
Dim lsCalificacion As String
Dim TOTALNORMME As Currency: Dim TOTALREFME As Currency: Dim TOTALMEN30ME As Currency
Dim TOTALMAY30ME  As Currency: Dim TOTALCOBJUDME  As Currency:  Dim TOTALLINCREDME As Currency
Dim TOTALCASTME As Currency
'------->
Dim gsNombreEntInfIBM As String
Dim gsNomCortoEntInfIBM As String
Dim gsCodEmpInfICC As String
Dim lsNumSecICC As String
'------->
Dim lsCadImp As String

Dim ofunc As New COMFunciones.FCOMCadenas
Dim ofuni As New COMFunciones.FCOMImpresion

Sql = "select nConsSisValor from ConstSistema where nConsSisCod=80"
Set rs = GetQuery(Sql)
gsCodEmpInfICC = IIf(IsNull(rs!nConsSisValor), "", rs!nConsSisValor)
gsNombreEntInfIBM = psNomCmact
gsNomCortoEntInfIBM = psCodCmac
lsPrefijos(1) = "DE LA"
lsPrefijos(2) = "DE LOS"
lsPrefijos(3) = "DE "
lsPrefijos(4) = "LA "
lsPrefijos(5) = "DEL "
lsPrefijos(6) = "DI "


' Crea el nombre de archivo.
'--COPIAR---
'lsNombreArchivo = App.Path & "\Spooler\ICS01.112"
'Open lsNombreArchivo For Output As #NúmeroArchivo

lsCadImp = lsCadImp & ofunc.FillNum(Trim(gsCodEmpInfICC), 3, " ") & ofunc.FillNum("1", 1, " ") & ofunc.FillNum(Trim(Format(dFecSis, "YYYYMMDD")), 8, " ") & _
        Space(11) & ofuni.ImpreFormat(gsNombreEntInfIBM, 80, 0) & ofuni.ImpreFormat(Trim(gsNomCortoEntInfIBM), 10, 0) & _
        Space(546) & "X"
'------>

Sql = "Select count(cCodPers) as Total From " & psServerConsol & "IBM1 where SUBSTRING(cApellPat,1,1) NOT IN('Ñ') "
Set rs = GetQuery(Sql)


contTotal = rs!Total
rs.Close
Set rs = Nothing

Sql = "Select cCodPers,cApellPat,cApellMat,cNombre,cTipPers,cTipDoc,cNumDoc,cCalifica," _
& " nNormMN,nRefMN,nVen30MN,nVen31MN,nCobJudMN,nLinCredMN,nCastMN,nNormME,nRefME, " _
& " nVen30ME , nVen31ME, nCobJudME, nLinCredME, nCastME " _
& " From " & psServerConsol & "IBM1 where SUBSTRING(cApellPat,1,1) <> 'Ñ' order by cApellPat,cApellmat "

'Select * From " & psServerConsol & "IBM1 where SUBSTRING(cApellPat,1,1) not in('Ñ') order by cApellPat,cApellmat"

Set rs = GetQuery(Sql)
Cont = 0
If Not (rs.EOF And rs.BOF) Then
    Do While Not rs.EOF
        Cont = Cont + 1
        lsNumSecICC = ofunc.FillNum(Trim(Str(Cont)), 6, "0")
        
        lsCalificacion = Trim(rs!cCalifica)
        'lsCalificacion = EmiteCalificacion(rs!cCodPers)
        'If Len(Trim(lsCalificacion)) = 0 Then
        '    lsCalificacion = Trim(rs!cCalifica)
        'End If
        
        lsCadImp = lsCadImp & ofunc.FillNum(Trim(gsCodEmpInfICC), 3, " ") & ofunc.FillNum(Trim("2"), 1, " ") & ofunc.FillNum(Trim(lsNumSecICC), 6, " ") & _
                    ofunc.FillNum(Trim(rs!CTIPDOC), 1, " ") & ofuni.ImpreFormat(Trim(rs!cNumDoc), 12, 0) & _
                    ofuni.ImpreFormat(Trim(Replace(rs!cApellPat, "Ñ", "#")), 60, 0) & ofuni.ImpreFormat(Trim(Replace(rs!cApellMat, "Ñ", "#")), 20, 0) & _
                    ofuni.ImpreFormat(Trim(Replace(rs!cNombre, "Ñ", "#")), 30, 0) & ofunc.FillNum(Trim(rs!cTipPers), 1, " ") & _
                    ofuni.ImpreFormat(lsCalificacion, 1, 0) & Space(5) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nNormMN), 0, rs!nNormMN)), 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nRefMN), 0, rs!nRefMN)), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen30MN), 0, rs!nVen30MN)), 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen31MN), 0, rs!nVen31MN)), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCobJudMN), 0, rs!nCobJudMN)), 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nLinCredMN), 0, rs!nLinCredMN)), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCastMN), 0, rs!nCastMN)), 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nNormME), 0, rs!nNormME)), 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nRefME), 0, rs!nRefME)), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen30ME), 0, rs!nVen30ME)), 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen31ME), 0, rs!nVen31ME)), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCobJudME), 0, rs!nCobJudME)), 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nLinCredME), 0, rs!nLinCredME)), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCastME), 0, rs!nCastME)), 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & Space(9) & "X" & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

Sql = "Select count(cCodPers) as Total From " & psServerConsol & "IBM1 where SUBSTRING(cApellPat,1,1) IN('Ñ') "
Set rs = GetQuery(Sql)

contTotal = rs!Total
rs.Close
Set rs = Nothing

Sql = "Select * From " & psServerConsol & "IBM1 where SUBSTRING(CAPELLPAT,1,1) in('Ñ') order by cApellPat,cApellmat"
Set rs = GetQuery(Sql)

m = 0
If Not (rs.EOF And rs.BOF) Then
    Do While Not rs.EOF
        m = m + 1
        Cont = Cont + 1
        lsNumSecICC = ofunc.FillNum(Trim(Str(Cont)), 6, "0")
        lsCalificacion = Trim(rs!cCalifica)
        'lsCalificacion = EmiteCalificacion(rs!cCodPers)
        'If Len(Trim(lsCalificacion)) = 0 Then
        '    lsCalificacion = Trim(rs!cCalifica)
        'End If
        lsCadImp = lsCadImp & ofunc.FillNum(Trim(gsCodEmpInfICC), 3, " ") & ofunc.FillNum(Trim("2"), 1, " ") & ofunc.FillNum(Trim(lsNumSecICC), 6, " ") & _
                    ofunc.FillNum(Trim(rs!CTIPDOC), 1, " ") & ofuni.ImpreFormat(Trim(rs!cNumDoc), 12, 0) & _
                    ofuni.ImpreFormat(Trim(Replace(rs!cApellPat, "Ñ", "#")), 60, 0) & ofuni.ImpreFormat(Trim(Replace(rs!cApellMat, "Ñ", "#")), 20, 0) & _
                    ofuni.ImpreFormat(Trim(Replace(rs!cNombre, "Ñ", "#")), 30, 0) & ofunc.FillNum(Trim(rs!cTipPers), 1, " ") & _
                    ofuni.ImpreFormat(lsCalificacion, 1, 0) & Space(5) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nNormMN), 0, rs!nNormMN)), 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nRefMN), 0, rs!nRefMN)), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen30MN), 0, rs!nVen30MN)), 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen31MN), 0, rs!nVen31MN)), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCobJudMN), 0, rs!nCobJudMN)), 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nLinCredMN), 0, rs!nLinCredMN)), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCastMN), 0, rs!nCastMN)), 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nNormME), 0, rs!nNormME)), 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nRefME), 0, rs!nRefME)), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen30ME), 0, rs!nVen30ME)), 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nVen31ME), 0, rs!nVen31ME)), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCobJudME), 0, rs!nCobJudME)), 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nLinCredME), 0, rs!nLinCredME)), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(IIf(IsNull(rs!nCastME), 0, rs!nCastME)), 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
                    ofuni.ImpreFormat(0, 15, 0, False) & Space(9) & "X" & Chr(10)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

Sql = " SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPPERS='1'"
Set rs1 = GetQuery(Sql)
TOTALPERSNAT = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPPERS='2'"
Set rs1 = GetQuery(Sql)
TOTALPERSJUR = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close


Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='1'"
Set rs1 = GetQuery(Sql)
TOTALTIPODOC1 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='2'"
Set rs1 = GetQuery(Sql)
TOTALTIPODOC2 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='3'"
Set rs1 = GetQuery(Sql)
TOTALTIPODOC3 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='4'"
Set rs1 = GetQuery(Sql)
TOTALTIPODOC4 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='5'"
Set rs1 = GetQuery(Sql)
TOTALTIPODOC5 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) as Total From " & psServerConsol & "IBM1 WHERE CTIPDOC='6'"
Set rs1 = GetQuery(Sql)
TOTALTIPODOC6 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
TOTALVENCMEN30MN = 0
Sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen30MN>0"
Set rs1 = GetQuery(Sql)
TOTALVENCMEN30MN = TOTALVENCMEN30MN + IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

TOTALVENCMEN30ME = 0
Sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen30ME>0"
Set rs1 = GetQuery(Sql)
TOTALVENCMEN30ME = TOTALVENCMEN30ME + IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

TOTALVENCMAY30MN = 0
Sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen31MN>0"
Set rs1 = GetQuery(Sql)
TOTALVENCMAY30MN = TOTALVENCMAY30MN + IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

TOTALVENCMAY30ME = 0
Sql = "SELECT COUNT(*) as total From " & psServerConsol & "IBM1 WHERE nVen31Me>0"
Set rs1 = GetQuery(Sql)
TOTALVENCMAY30ME = TOTALVENCMAY30ME + IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

Sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='0'"
Set rs1 = GetQuery(Sql)
TOTALCAL0 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='1'"
Set rs1 = GetQuery(Sql)
TOTALCAL1 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='2'"
Set rs1 = GetQuery(Sql)
TOTALCAL2 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='3'"
Set rs1 = GetQuery(Sql)
TOTALCAL3 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close
Sql = "SELECT COUNT(*) AS TOTAL From " & psServerConsol & "IBM1 WHERE CCALIFICA ='4'"
Set rs1 = GetQuery(Sql)
TOTALCAL4 = IIf(IsNull(rs1!Total), 0, rs1!Total)
rs1.Close

Sql = "SELECT SUM(nNormMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALNORMMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nRefMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALREFMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nVen30MN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALMEN30MN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nVen31MN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALMAY30MN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nCobJudMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALCOBJUDMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nLinCredMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALLINCREDMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nCastMN) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALCASTMN = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close

'monto soles
Sql = "SELECT SUM(nNormME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALNORMME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nRefME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALREFME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nVen30ME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALMEN30ME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nVen31ME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALMAY30ME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nCobJudME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALCOBJUDME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nLinCredME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALLINCREDME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close
Sql = "SELECT SUM(nCastME) AS TOTAL From " & psServerConsol & "IBM1"
Set rs1 = GetQuery(Sql)
TOTALCASTME = IIf(IsNull(rs1!Total), 0, Format(rs1!Total, "#0.00"))
rs1.Close

lsCadImp = lsCadImp & ofunc.FillNum(Trim(gsCodEmpInfICC), 3, " ") & ofunc.FillNum(Trim("3"), 1, " ") & _
ofuni.ImpreFormat(TOTALTIPODOC1, 8, 0) & ofuni.ImpreFormat(TOTALTIPODOC2, 8, 0) & ofuni.ImpreFormat(TOTALTIPODOC3, 8, 0) & ofuni.ImpreFormat(TOTALTIPODOC4, 8, 0) & ofuni.ImpreFormat(TOTALTIPODOC5, 8, 0) & ofuni.ImpreFormat(TOTALTIPODOC6, 8, 0) & _
ofuni.ImpreFormat(TOTALPERSNAT, 8, 0) & ofuni.ImpreFormat(TOTALPERSJUR, 8, 0) & _
ofuni.ImpreFormat(TOTALVENCMEN30MN, 8, 0) & ofuni.ImpreFormat(TOTALVENCMAY30MN, 8, 0) & ofuni.ImpreFormat(TOTALVENCMEN30ME, 8, 0) & ofuni.ImpreFormat(TOTALVENCMAY30ME, 8, 0) & _
ofuni.ImpreFormat(TOTALCAL0, 8, 0) & ofuni.ImpreFormat(TOTALCAL1, 8, 0) & ofuni.ImpreFormat(TOTALCAL2, 8, 0) & ofuni.ImpreFormat(TOTALCAL3, 8, 0) & ofuni.ImpreFormat(TOTALCAL4, 8, 0) & _
ofuni.ImpreFormat(EliminaPunto(TOTALNORMMN), 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(TOTALREFMN), 15, 0, False) & _
ofuni.ImpreFormat(EliminaPunto(TOTALMEN30MN), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(TOTALMAY30MN), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(TOTALCOBJUDMN), 15, 0, False) & _
ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(TOTALLINCREDMN), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(TOTALCASTMN), 15, 0, False) & _
ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
ofuni.ImpreFormat(0, 15, 0, False) & _
ofuni.ImpreFormat(EliminaPunto(TOTALNORMME), 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(TOTALREFME), 15, 0, False) & _
ofuni.ImpreFormat(EliminaPunto(TOTALMEN30ME), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(TOTALMAY30ME), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(TOTALCOBJUDME), 15, 0, False) & _
ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(TOTALLINCREDME), 15, 0, False) & ofuni.ImpreFormat(EliminaPunto(TOTALCASTME), 15, 0, False) & _
ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & ofuni.ImpreFormat(0, 15, 0, False) & _
ofuni.ImpreFormat(0, 15, 0, False) & Space(9) & "X" & Chr(10)

Set ofunc = Nothing
Set ofuni = Nothing
End Sub

Private Sub Class_Initialize()
Set Conecta = New comconecta.DCOMConecta
End Sub

Private Sub Class_Terminate()
Set Conecta = Nothing
End Sub

Private Function ValidaTabla(pdFecSis As Date) As Boolean
Dim Sql As String
Dim rs As Recordset

Sql = "SELECT * FROM sysobjects WHERE name = '" & "RCDvc" & Format(pdFecSis, "yyyymm") & "01'"
Conecta.AbreConexion
Set rs = Conecta.CargaRecordSet(Sql)
ValidaTabla = True

If rs.EOF And rs.BOF Then
    Conecta.CierraConexion
    MsgBox "Tabla [RCDvc" & Format(pdFecSis, "yyyymm") & "01]  No existe, Por favor comuniquese con el Dpto de Sistemas", vbInformation, "Aviso"
    ValidaTabla = False
End If

End Function

Public Function ObtenerNuevoCredCapArmotizado(ByVal sFechaIni As String, ByVal i As Integer) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = " select isnull(sum(nNumDesembN),0) NroDesembN ,"
    Sql = Sql & "       isnull(sum(nMontoDesembN) ,0) MontoDesembN,"
    Sql = Sql & "       isnull(sum(nCapPag),0) CapPag"
    Sql = Sql & " From ColocEstadDiaCred"
    Sql = Sql & " where datediff(m,dEstad,'" & Format(DateAdd("m", i, sFechaIni), "yyyy/mm/dd") & "')=0"
    Sql = Sql & " and cLineaCred='02'"
    rs.CursorLocation = adUseClient
    Set rs = GetQuery(Sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerNuevoCredCapArmotizado = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerNuevoCredCapArmotizado = Nothing
    End If
End Function

Public Function ObtenerMontoPrestamos(ByVal sFechaIni As String, ByVal i As Integer) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = "       select isnull(sum(nNumCanc),0) NroCanc ,"
    Sql = Sql & "        isnull(sum(nSaldoCap) ,0) SaldoCap"
    Sql = Sql & " From ColocEstadDiaCred"
    Sql = Sql & " where datediff(m,dEstad,'" & Format(DateAdd("m", i, sFechaIni), "yyyy/mm/dd") & "')=0"
    Sql = Sql & " and cLineaCred='02'"""
    rs.CursorLocation = adUseClient
    Set rs = GetQuery(Sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerMontoPrestamos = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerMontoPrestamos = Nothing
    End If
    
End Function

Public Function ObtenerCarteraMoroso(ByVal sFechaIni As String, ByVal i As Integer) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = " select isnull(sum(nNumJud),0) NroJudi , isnull(sum(nMontoJudic) ,0) MontoJud "
    Sql = Sql & " From ColocEstadDiaCred "
    Sql = Sql & " where datediff(m,dEstad,'" & Format(DateAdd("m", i, sFechaIni), "yyyy/mm/dd") & "')=0 "
    Sql = Sql & " and cLineaCred='02'"
    rs.CursorLocation = adUseClient
    Set rs = GetQuery(Sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerCarteraMoroso = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerCarteraMoroso = Nothing
    End If
End Function

Public Function ObtenerCarteraRefinanciada(ByVal sFechaIni As String, ByVal i As Integer) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = "       Select isnull(sum(nNumRef),0) NroRef ,"
    Sql = Sql & "       isnull(sum(nMontoRefinan) ,0) MonoRefinan,"
    Sql = Sql & "       isnull(sum(nIntPag),0) nIntPag,"
    Sql = Sql & "       isnull(sum(nCapPag),0) nCapPag,"
    Sql = Sql & "       isnull(sum(nNumCanc),0) NroCan,"
    Sql = Sql & "       isnull(sum(nSaldoCap),0) SaldoCap"
    Sql = Sql & " From ColocEstadDiaCred"
    Sql = Sql & " where datediff(m,dEstad,'" & Format(DateAdd("m", i, sFechaIni), "yyyy/mm/dd") & "')=0"
    Sql = Sql & " and cLineaCred='02'"
    rs.CursorLocation = adUseClient
    Set rs = GetQuery(Sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerCarteraRefinanciada = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerCarteraRefinanciada = Nothing
    End If
End Function

Public Function ObtenerSolicitudesPresentadas(ByVal sFechaIni As String, ByVal pdFechaCierre As Date, ByVal psServConsol As String) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = Sql & "SELECT Fecha, Count(cCtaCod) Numero, Sum(Monto) Monto from ("
    Sql = Sql & " Select Convert(Varchar(7),dAsignacion,102) Fecha, cCtaCod, Monto = CASE"
    Sql = Sql & "    WHEN nPrdEstado In ('2000','2003') THEN nMontoSol"
    Sql = Sql & "    WHEN nPrdEstado In ('2001') THEN nMontoSug"
    Sql = Sql & "    WHEN nPrdEstado In ('2002','2201','2080') THEN nMontoApr"
    Sql = Sql & "    WHEN nPrdEstado In ('2020','2021','2022','2030','2031','2032','2050') THEN nMontoDesemb END"
    Sql = Sql & " From " & psServConsol & "CreditoConsolTotal"
    Sql = Sql & " where SubString(cLineaCred,1,2) = '02'"
    Sql = Sql & " And nPrdEstado In ('2000','2003','2001','2002','2201','2080','2020','2021','2022','2030','2031','2032','2050')"
    Sql = Sql & " and DateDiff(dd,dAsignacion,'" & Format(sFechaIni, "mm/dd/yyyy") & "') <= 0  And DateDiff(dd,dAsignacion,'" & Format(pdFechaCierre, "mm/dd/yyyy") & "') >= 0"
    Sql = Sql & " ) C Group by Fecha Order by Fecha"
    rs.CursorLocation = adUseClient
    Set rs = GetQuery(Sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerSolicitudesPresentadas = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerSolicitudesPresentadas = Nothing
    End If
End Function

Public Function ObtenerCreditosAprobados(ByVal sFechaIni As String, ByVal pdFechaCierre As Date, ByVal psServConsol As String) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = "SELECT Fecha, Count(cCtaCod) Numero, Sum(Monto) Monto from ("
    Sql = Sql & " Select Convert(Varchar(7),dFecApr,102) Fecha, cCtaCod, Monto = CASE"
    Sql = Sql & "     WHEN nPrdEstado In ('2002') THEN nMontoApr"
    Sql = Sql & "     WHEN nPrdEstado In ('2050','2020','2021','2022','2030','2031','2032') THEN nMontoDesemb END"
    Sql = Sql & " From " & psServConsol & "CreditoConsolTotal"
    Sql = Sql & " where SubString(cLineaCred,1,2) = '02'"
    Sql = Sql & " And nPrdEstado In ('2020','2021','2022','2050','2030','2031','2032','2002')"
    Sql = Sql & " and DateDiff(dd,dFecApr,'" & Format(sFechaIni, "mm/dd/yyyy") & "') <= 0  And DateDiff(dd,dFecApr,'" & Format(pdFechaCierre, "mm/dd/yyyy") & "') >= 0"
    Sql = Sql & " ) C Group by Fecha Order by Fecha"
    rs.CursorLocation = adUseClient
    Set rs = GetQuery(Sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerCreditosAprobados = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerCreditosAprobados = Nothing
    End If

End Function

Public Function ObtenerCreditosDesembolsados(ByVal sFechaIni As String, ByVal pdFechaCierre As Date, ByVal psServConsol As String) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = " Select Convert(Varchar(7),dFecVig,102) Fecha, Count(cCtaCod) Numero, Sum(nMontoDesemb) Monto"
    Sql = Sql & " From " & psServConsol & "CreditoConsolTotal where SubString(cLineaCred,1,2) = '02' And nPrdEstado In ('2201','2020','2021','2022','2050','2031','2032','2002')"
    Sql = Sql & " and DateDiff(dd,dFecVig,'" & Format(sFechaIni, "mm/dd/yyyy") & "') <= 0  And DateDiff(dd,dFecVig,'" & Format(pdFechaCierre, "mm/dd/yyyy") & "') >= 0"
    Sql = Sql & " Group by Convert(Varchar(7),dFecVig,102) Order by Fecha"
    rs.CursorLocation = adUseClient
    Set rs = GetQuery(Sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerCreditosDesembolsados = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerCreditosDesembolsados = Nothing
    End If
End Function

Public Function ObtenerPlazos(ByVal sFechaIni As String, ByVal psServConsol As String, ByVal i As Integer) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = " Select Rango, Count(Cuenta) Num, isnull(Sum(Desembolso),0) Desembolso, Sum(Saldo) Saldo FROM ("
    Sql = Sql & " Select C.cuenta, C.Desembolso, C.Saldo, C.Plazo , Rango = CASE"
    Sql = Sql & "     WHEN C.Plazo <= 180 THEN '01. Hasta 6 Meses'"
    Sql = Sql & "     WHEN C.Plazo > 180 And Plazo <= 360 THEN '02. De 6 a 12 Meses'"
    Sql = Sql & "     WHEN C.Plazo > 360 And Plazo <= 720  THEN '03. De 12 a 24 Meses'"
    Sql = Sql & "     WHEN C.Plazo > 720  THEN '04. Mas de 24 Meses' END FROM ("
    Sql = Sql & " Select cCtaCod cuenta, nMontoDesemb Desembolso, nsaldocap saldo, ncuotasapr*nplazoapr plazo"
    Sql = Sql & " FROM " & psServConsol & "CreditoConsolTotal WHERE nPrdEstado IN ('2020','2021','2022','2050','2030','2031','2032','2002') And SubString(cLineaCred,1,2) = '02'"
    Sql = Sql & "  and nTipcuota in ('10','11','20','21','30','31')"
    Sql = Sql & "  and DateDiff(mm,dFecVig,'" & Format(DateAdd("m", i, sFechaIni), "yyyy/mm/dd") & "') = 0"
    Sql = Sql & " Union"
    Sql = Sql & " Select cCtaCod cuenta, nMontoDesemb Desembolso, nsaldocap saldo, ncuotasapr*30 plazo"
    Sql = Sql & " FROM " & psServConsol & "CreditoConsolTotal WHERE nPrdEstado IN ('2020','2021','2022','2050','2030','2031','2032','2002') and SubString(cLineaCred,1,2) = '02'"
    Sql = Sql & "  and nTipcuota in ('40','41','50','51','60','61','70','71')"
    Sql = Sql & "  and DateDiff(mm,dFecVig,'" & Format(DateAdd("m", i, sFechaIni), "yyyy/mm/dd") & "') = 0"
    Sql = Sql & " ) C"
    Sql = Sql & " ) C1 Group by Rango Order by Rango"
    rs.CursorLocation = adUseClient
    Set rs = GetQuery(Sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerPlazos = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerPlazos = Nothing
    End If
End Function

Public Function ObtenerPlazoporSexo(ByVal sFechaIni As String, ByVal psServConsol As String, ByVal i As Integer) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = " SELECT cSexPers, count(cCtaCod) Cuenta , isnull(sum(T.Desembolso),0) Desembolso FROM ("
    Sql = Sql & "    Select"
    Sql = Sql & "    cSexPers = CASE WHEN cPersNatSexo = 'M' THEN 'H'"
    Sql = Sql & "              WHEN cPersNatSexo = 'F' THEN 'M' ELSE 'J'"
    Sql = Sql & "               END,"
    Sql = Sql & "    C.cCtaCod , nMontoDesemb Desembolso"
    Sql = Sql & "    From " & psServConsol & "CreditoConsolTotal c"
    Sql = Sql & "    Inner Join " & psServConsol & "ProductoPersonaConsol PE on C.cCtaCod = Pe.cCtaCod"
    Sql = Sql & "    Inner Join Persona P on P.cPersCod = Pe.cPersCod"
    Sql = Sql & "    left outer join PersonaNat PN on PN.cPersCod = P.cPersCod"
    Sql = Sql & "    Where "
    Sql = Sql & "    C.nPrdEstado in ('2020','2021','2022','2050','2030','2031','2032') and substring(cLineaCred,1,2) = '02'"
    Sql = Sql & "    and PE.nPrdPersRelac = 20"
    Sql = Sql & "   and DateDiff(mm,dfecvig,'" & Format(DateAdd("m", i, sFechaIni), "yyyy/mm/dd") & "') = 0"
    Sql = Sql & " ) T"
    Sql = Sql & " GROUP BY T.cSexPers ORDER BY T.cSexPers"
    rs.CursorLocation = adUseClient
    Set rs = GetQuery(Sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerPlazoporSexo = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerPlazoporSexo = Nothing
    End If
End Function

Public Function ObtenerColocacionesporTrabajador(ByVal sFechaIni As String, ByVal psServConsol As String, ByVal i As Integer) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = " Select Fecha, NroTraba, Count(Cuenta) Num, isnull(Sum(Desembolso),0) Desembolso, Sum(Saldo) Saldo FROM"
    Sql = Sql & " (  Select Fecha, C.cuenta, C.Desembolso, C.Saldo, NroTRaba = CASE"
    Sql = Sql & "   WHEN nNroEmp <= 10 THEN '01. Hasta 10'"
    Sql = Sql & "   WHEN nNroEmp > 10 THEN '02. Más de 10' END"
    Sql = Sql & " From"
    Sql = Sql & " ("
    Sql = Sql & "   Select Convert(Varchar(7),dfecvig,102) Fecha, cCtaCod cuenta, nMontoDesemb Desembolso, nsaldocap saldo, ISNULL(F.nnroemp,0) nnroemp"
    Sql = Sql & "        FROM " & psServConsol & "CreditoConsolTotal X"
    Sql = Sql & "        inner join   " & psServConsol & "FuenteingresoConsol F on X.cNumFuente = F.cNumFuente"
    Sql = Sql & "     WHERE nPrdEstado in ('2020','2021','2022','2050','2030','2031','2032') and substring(cLineaCred,1,2) = '02'"
    Sql = Sql & "       and DateDiff(mm,dfecvig,'" & Format(DateAdd("m", i, sFechaIni), "yyyy/mm/dd") & "')=0 ) C "  '= 0 And DateDiff(dd,dfecvig,'" & Format(pdFechaCierre, "mm/dd/yyyy") & "') >= 0 ) C"
    Sql = Sql & " ) C1 Group by Fecha, NroTRaba Order by Fecha, NroTRaba"
    rs.CursorLocation = adUseClient
    Set rs = GetQuery(Sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerColocacionesporTrabajador = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerColocacionesporTrabajador = Nothing
    End If
End Function

Public Function ObtenerColocacionesporMonto(ByVal sFechaIni As String, ByVal psServConsol As String, ByVal i As Integer) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = " Select  Convert(Varchar(7),dfecvig,102) Fecha,"
    Sql = Sql & " cDestCre = CASE WHEN nDestCre = '2' THEN 'A' WHEN nDestCre = '1' THEN 'B'"
    Sql = Sql & " WHEN nDestCre = '3' THEN 'C' END,"
    Sql = Sql & " COUNT(C.cCtaCod) numero, sum(nMontoDesemb) Desembolso"
    Sql = Sql & "       FROM " & psServConsol & "CreditoConsolTotal C"
    Sql = Sql & " WHERE nPrdEstado in ('2020','2021','2022','2050','2030','2031','2032') and substring(cLineaCred,1,2) = '02'"
    Sql = Sql & "     And DateDiff(dd,dFecVig,'" & Format(DateAdd("m", i, sFechaIni), "yyyy/mm/dd") & "') = 0 " 'And DateDiff(dd,dfecvig,'" & Format(pdFechaCierre, "mm/dd/yyyy") & "') >= 0"
    Sql = Sql & " GROUP BY Convert(Varchar(7),dfecvig,102), nDestCre ORDER BY Fecha, nDestCre"
    rs.CursorLocation = adUseClient
    Set rs = GetQuery(Sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerColocacionesporMonto = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerColocacionesporMonto = Nothing
    End If
    
End Function

Public Function ObtenerColocacinesporMonto1(ByVal sFechaIni As String, ByVal psServConsol As String, ByVal i As Integer) As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = "Select Fecha, Rango, Count(Cuenta) Num, isnull(Sum(Desembolso),0) Desembolso, Sum(Saldo) Saldo FROM"
    Sql = Sql & " (  Select Fecha, C.cuenta, C.Desembolso, C.Saldo, Rango = CASE"
    Sql = Sql & "   WHEN C.Desembolso <= 2000 THEN '01. Hasta 2000 NS'"
    Sql = Sql & "   WHEN C.Desembolso > 2000 AND C.Desembolso <= 5000 THEN '02. De 2001 A 5000'"
    Sql = Sql & "   WHEN C.Desembolso > 5000 AND C.Desembolso <= 10000 THEN '03. De 5001 a 10000'"
    Sql = Sql & "   WHEN C.Desembolso > 10000  THEN '04. Mas de 10000' END"
    Sql = Sql & " From"
    Sql = Sql & " ("
    Sql = Sql & "   Select Convert(Varchar(7),dfecvig,102) Fecha, cCtaCod cuenta, nMontoDesemb Desembolso, nsaldocap saldo"
    Sql = Sql & "        FROM " & psServConsol & "CreditoConsolTotal WHERE nPrdEstado in ('2020','2021','2022','2050','2030','2031','2032') and substring(cLineaCred,1,2) = '02'"
    Sql = Sql & "       And DateDiff(dd,dFecVig,'" & Format(DateAdd("m", i, sFechaIni), "yyyy/mm/dd") & "') = 0) C " ' And DateDiff(dd,dfecvig,'2003/03/30') >= 0 ) C
    Sql = Sql & " ) C1 Group by Fecha, Rango Order by Fecha, Rango"
    rs.CursorLocation = adUseClient
    Set rs = GetQuery(Sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerColocacinesporMonto1 = rs
        Set rs.ActiveConnection = Nothing
    Else
        Set ObtenerColocacinesporMonto1 = Nothing
    End If
End Function

