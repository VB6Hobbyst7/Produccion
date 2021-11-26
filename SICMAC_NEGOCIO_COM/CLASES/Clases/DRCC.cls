VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True



Sub BorraTablaRCC(ByVal psServer As String)
Dim Sql1 As String
Dim Sql2 As String

Dim Co As DConecta

Set Co = New DConecta

Sql1 = "delete " & psServer & "rcctotal"
Sql2 = "delete " & psServer & "rcctotaldet"


On Error GoTo RCC
    Co.AbreConexion
    Co.Ejecutar (Sql1)
    Co.Ejecutar (Sql2)
    Co.CierraConexion
    Set Co = Nothing
Exit Sub
RCC:
    Set Co = Nothing
   Call RaiseError(MyUnhandledError, "DRCC:BorraTablaRCC Method")
End Sub

Sub InsertaRccHistorico(ByVal pcConexServerConsol As ADODB.Connection)
Dim Sql1 As String
Dim Sql2 As String
Dim Co As DConecta

Set Co = New DConecta

'Sql1 = " Insert " & psServer & "rcchistotal"
'Sql1 = Sql1 & " Select * from " & psServer & "rcctotal"

'Sql2 = "Insert " & psServer & "rcchistotaldet"
'Sql2 = Sql2 & " Select * from " & psServer & "rcctotaldet"

Sql1 = " Insert rcchistotal"
Sql1 = Sql1 & " Select * from rcctotal"

Sql2 = "Insert rcchistotaldet"
Sql2 = Sql2 & " Select * from rcctotaldet"

pcConexServerConsol.CommandTimeout = 12000
pcConexServerConsol.Execute Sql1
pcConexServerConsol.Execute Sql2

'On Error GoTo RCC
'    Co.AbreConexion
'    Co.ConexionActiva.CommandTimeout = 12000
'    Co.Ejecutar (Sql1)
'    Co.Ejecutar (Sql2)
'    Co.CierraConexion
'    Set Co = Nothing
Exit Sub
RCC:
   Set Co = Nothing
   Call RaiseError(MyUnhandledError, "DRCC:InsertaRCChistorico Method")
End Sub


Function GetFecha(ByVal psServer As String, ByVal nTpo As Integer) As String
'Definicion
'1 RccTotal
'2 RccTotalDet
'3 RccHisTotal
'4 RccHisTotalDet
Dim Sql As String, lsMensaje As String
Dim Con As DConecta
Dim Rs As ADODB.Recordset
Dim sFecha As String

Select Case nTpo
    Case 1
        Sql = "select max(Fec_Rep) Fecha from " & psServer & "RccTotal"
        lsMensaje = " en RccTotal "
    Case 2
        Sql = "select max(dFecha) Fecha from " & psServer & "RccTotalDet"
        lsMensaje = " en RccTotalDet "
    Case 3
        Sql = "select max(Fec_Rep) Fecha from " & psServer & "RccHisTotal "
        lsMensaje = " en RccHisTotal "
    Case 4
        Sql = "select max(dFecha) Fecha from " & psServer & "RccHisTotalDet"
        lsMensaje = " en RccHisTotalDet "
End Select

Set Con = New DConecta

Con.AbreConexion
Set Rs = Con.CargaRecordSet(Sql)
If Not (Rs.EOF And Rs.BOF) Then
    sFecha = IIf(IsNull(Rs!Fecha), "-1", Rs!Fecha)
Else
    sFecha = "-1"
End If

Con.CierraConexion

If sFecha = "-1" Then
    MsgBox "No exiten datos " & lsMensaje, vbCritical, "AVISO"
End If

GetFecha = Format(sFecha, "DD/MM/YYYY")
Set Rs = Nothing
Set Co = Nothing
End Function

Sub InsertaRccCab(ByVal psServer As String, ByVal sFecCD As String)
Dim Sql As String
Dim Co As DConecta

Set Co = New DConecta

Sql = " Insert " & psServer & "rcctotal"

Sql = Sql & " Select"
Sql = Sql & " convert(datetime,substring(Col001,12,8)) Fec_Rep,"
Sql = Sql & " substring(Col001,2,10) Cod_Edu, null Tip_Det,"
Sql = Sql & " substring(Col001,20,1) Tip_Doc_Trib,"
Sql = Sql & " substring(Col001,21,11) Cod_Doc_Trib,"
Sql = Sql & " substring(Col001,32,1) Tip_Doc_Id,"
Sql = Sql & " substring(Col001,33,12) Cod_Doc_Id,"
Sql = Sql & " substring(Col001,45,1) TipPers,"
Sql = Sql & " substring(Col001,46,1) TipEmp,"
Sql = Sql & " convert(int,substring(Col001,47,3)) Can_ents,"
Sql = Sql & " isnull(convert(real,substring(Col001,50,5)/100),'') Calif_0,"
Sql = Sql & " isnull(convert(real,substring(Col001,55,5)/100),'') Calif_1,"
Sql = Sql & " isnull(convert(real,substring(Col001,60,5)/100),'') Calif_2,"
Sql = Sql & " isnull(convert(real,substring(Col001,65,5)/100),'') Calif_3,"
Sql = Sql & " isnull(convert(real,substring(Col001,70,5)/100),'') Calif_4,"
Sql = Sql & " case when len(rtrim(ltrim(substring(Col001,75,120))))=0 then '' else rtrim(ltrim(substring(Col001,75,120))) + ' ' end +"
Sql = Sql & " case when len(IsNull(rtrim(ltrim(substring(Col001,195,40))),'')) = 0 then '' else  + IsNull(rtrim(ltrim(substring(Col001,195,40))),'')  + ' ' end +"
Sql = Sql & " case when len(rtrim(ltrim(substring(Col001,235,40))))=0 then '' else rtrim(ltrim(substring(Col001,235,40))) + ' ' end +"
Sql = Sql & " case when len(rtrim(ltrim(substring(Col001,275,40))))=0 then '' else rtrim(ltrim(substring(Col001,275,40))) + ' ' end +"
Sql = Sql & " case when len(rtrim(ltrim(substring(Col001,315,40))))=0 then '' else rtrim(ltrim(substring(Col001,315,40))) end Nom_Deu"
Sql = Sql & " From " & psServer & "rcc" & sFecCD
Sql = Sql & " where substring(Col001,1,1)='1'"


On Error GoTo RccErr
    Co.AbreConexion
    'Co.AdicionaCmdBatch (Sql)
    Co.ConexionActiva.CommandTimeout = 12000
    Co.Ejecutar (Sql)
    
    Co.CierraConexion
    Set Co = Nothing
Exit Sub

RccErr:
   Call RaiseError(MyUnhandledError, "DRcc:InsertaRccDet Method")
   Set Co = Nothing
End Sub

Sub InsertRCCDet(ByVal psServer As String, ByVal sFecCD As String)
Dim Sql As String
Dim sF As String
Dim Co As DConecta

Set Co = New DConecta
'Calcula ultimo dia del mes
sF = "01/" & Right(sFecCD, 2) & "/" & Left(sFecCD, 4)
sF = DateAdd("M", 1, sF)
sF = DateAdd("D", -1, sF)

Sql = "Insert " & psServer & "rcctotaldet"
Sql = Sql & "(dFecha,Cod_Edu,Cod_Emp,Tip_Credito, Cod_Cuenta,Condicion,Val_Saldo,Clasificacion)"

'Sql = Sql & " Select   convert(datetime,'" & Format(sF, "YYYY/MM/DD") & "'),substring(Col001,3,10)codedu,"
'Sql = Sql & " isnull(substring(Col001,2,1),'')tipdet,isnull(substring(Col001,13,5),'')codemp,"
'Sql = Sql & " isnull(substring(Col001,18,1),'')tipcredito,"
'Sql = Sql & " isnull(substring(Col001,19,14),'')codcuenta,"
'Sql = Sql & " case when isnumeric(substring(Col001,33,4)) = 1 then substring(Col001,33,4) else  null end condicion,"
'Sql = Sql & " convert(money,substring(Col001,37,15))/100 valsaldo,"
'Sql = Sql & " isnull(substring(Col001,52,1),'')clasificacion"
'Sql = Sql & " From " & psServer & "rcc" & sFecCD
'Sql = Sql & " where substring(Col001,1,1)='2'"


Sql = Sql & " Select"
Sql = Sql & " convert(datetime,'" & Format(sF, "YYYY/MM/DD") & "') dFecha,"
Sql = Sql & " substring(Col001,2,10) Cod_Edu,"
Sql = Sql & " substring(Col001,12,5) Cod_Emp,"
Sql = Sql & " substring(Col001,17,1) Tip_Credito,"
Sql = Sql & " isnull(substring(Col001,18,14),'') Cod_Cuenta,"
Sql = Sql & " case when isnumeric(substring(Col001,32,4)) = 1 then substring(Col001,32,4) else  null end Condicion,"
Sql = Sql & " convert(money,substring(Col001,36,15))/100 Val_Saldo,"
Sql = Sql & " substring(Col001,51,1) Clasificacion"
Sql = Sql & " From " & psServer & "rcc" & sFecCD
Sql = Sql & " where substring(Col001,1,1)='2'"


On Error GoTo RccErr
    Co.AbreConexion
    Co.ConexionActiva.CommandTimeout = 12000
    Co.Ejecutar (Sql)
    Co.CierraConexion
    Set Co = Nothing
Exit Sub

RccErr:
   Call RaiseError(MyUnhandledError, "DRcc:InsertaRccDet Method")
   Set Co = Nothing
End Sub

Sub InsertaDataColocCalifProvTotal(ByVal sServer As String, ByVal psFecData As String)
Dim Sql As String
Dim Co As DConecta
Set Co = New DConecta
Sql = " INSERT " & sServer & "ColocCalifProvTotal "
Sql = Sql & " SELECT '" & Format(psFecData, "YYYY/MM/DD") & "', * FROM ColocCalifProv"
Co.AbreConexion
Co.Ejecutar (Sql)
Co.CierraConexion
Set Co = Nothing
End Sub

Sub ActulizaFechaCierre(ByVal psFecData As String)
Dim Sql As String
Dim Co As DConecta
Set Co = New DConecta
Co.AbreConexion
Sql = "UPDATE CONSTSISTEMA SET NCONSSISVALOR= '" & psFecData & "' where nConsSisCod = 160"
Co.Ejecutar (Sql)
Co.CierraConexion
Set Co = Nothing
End Sub
Function VerificaDataMigradaFecha(ByVal psServer As String, ByVal psFecData As String) As Long
Dim Sql As String
Dim Co As DConecta
Dim Rs As ADODB.Recordset

Set Co = New DConecta

Sql = " select count(*) nro from " & psServer & "ColocCalifProvTotal "
Sql = Sql & " where datediff(day,dfecha,'" & Format(psFecData, "YYYY/MM/DD") & "')=0"

Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
VerificaDataMigradaFecha = Rs!Nro
If Rs!Nro > 0 Then
    MsgBox "La Data ya fue Transferida", vbInformation, "AVISO"
    Exit Function
End If
Set Co = Nothing
Set Rs = Nothing
End Function
Function FechasRcc(ByVal psServer As String) As ADODB.Recordset
Dim Sql As String
Dim Co As DConecta
Dim Rs As ADODB.Recordset

Set Co = New DConecta

Sql = "select distinct fec_Rep Fecha from " & psServer & "rcchistotal"
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
Co.CierraConexion
Set FechasRcc = Rs

Set Co = Nothing
Set Rs = Nothing
End Function

Function ServConsol(ByVal pnServer As Integer) As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Sql = "Select nConsSisValor serv from ConstSistema where nConsSisCod = " & pnServer
Set Co = New DConecta
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
If Not (Rs.EOF And Rs.BOF) Then
    ServConsol = Trim(Rs!Serv)
End If
Co.CierraConexion
Set Rs = Nothing
Set Co = Nothing
End Function

Function NombreServerConsol() As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta
Sql = "SELECT Substring(nConsSisValor,1,LEN(nConsSisValor) - 5) SERVIDOR FROM ConstSistema WHERE nConsSisCod  = 43"
Co.AbreConexion
Set Rs = Co.CargaRecordSet(Sql)
If Not (Rs.EOF And Rs.BOF) Then
    NombreServerConsol = Trim(Rs!SERVIDOR)
End If
Co.CierraConexion
Set Rs = Nothing
Set Co = Nothing
End Function

Function VerificaTabla(ByVal psServer As String, ByVal sFecCD As String) As Integer
Dim Sql As String
Dim Co As DConecta
Dim Rs As ADODB.Recordset

Sql = " select * from " & psServer & "sysobjects  where name = 'Rcc" & sFecCD & "'"
Set Co = New DConecta
Co.AbreConexion


Set Rs = Co.CargaRecordSet(Sql)
If Not (Rs.EOF And Rs.BOF) Then
    VerificaTabla = 1
Else
    VerificaTabla = -1
End If
Co.CierraConexion
Set Co = Nothing
Set Rs = Nothing
End Function


