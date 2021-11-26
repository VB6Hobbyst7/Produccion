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
Dim SQL1 As String
Dim Sql2 As String

Dim Co As DConecta

Set Co = New DConecta

SQL1 = "delete " & psServer & "rcctotal"
Sql2 = "delete " & psServer & "rcctotaldet"


On Error GoTo RCC
    Co.AbreConexion
    Co.Ejecutar (SQL1)
    Co.Ejecutar (Sql2)
    Co.CierraConexion
    Set Co = Nothing
Exit Sub
RCC:
    Set Co = Nothing
   Call RaiseError(MyUnhandledError, "DRCC:BorraTablaRCC Method")
End Sub

Sub InsertaRccHistorico(ByVal pcConexServerConsol As ADODB.Connection)
Dim SQL1 As String
Dim Sql2 As String
Dim Co As DConecta

Set Co = New DConecta

'Sql1 = " Insert " & psServer & "rcchistotal"
'Sql1 = Sql1 & " Select * from " & psServer & "rcctotal"

'Sql2 = "Insert " & psServer & "rcchistotaldet"
'Sql2 = Sql2 & " Select * from " & psServer & "rcctotaldet"

SQL1 = " Insert rcchistotal"
SQL1 = SQL1 & " Select * from rcctotal"

Sql2 = "Insert rcchistotaldet"
Sql2 = Sql2 & " Select * from rcctotaldet"

pcConexServerConsol.CommandTimeout = 12000
pcConexServerConsol.Execute SQL1
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
Dim sql As String, lsmensaje As String
Dim Con As DConecta
Dim rs As ADODB.Recordset
Dim sFecha As String

Select Case nTpo
    Case 1
        sql = "select max(Fec_Rep) Fecha from " & psServer & "RccTotal"
        lsmensaje = " en RccTotal "
    Case 2
        sql = "select max(dFecha) Fecha from " & psServer & "RccTotalDet"
        lsmensaje = " en RccTotalDet "
    Case 3
        sql = "select max(Fec_Rep) Fecha from " & psServer & "RccHisTotal "
        lsmensaje = " en RccHisTotal "
    Case 4
        sql = "select max(dFecha) Fecha from " & psServer & "RccHisTotalDet"
        lsmensaje = " en RccHisTotalDet "
End Select

Set Con = New DConecta

Con.AbreConexion
Set rs = Con.CargaRecordSet(sql)
If Not (rs.EOF And rs.BOF) Then
    sFecha = IIf(IsNull(rs!Fecha), "-1", rs!Fecha)
Else
    sFecha = "-1"
End If

Con.CierraConexion

If sFecha = "-1" Then
    MsgBox "No exiten datos " & lsmensaje, vbCritical, "AVISO"
End If

GetFecha = Format(sFecha, "DD/MM/YYYY")
Set rs = Nothing
Set Co = Nothing
End Function

Sub InsertaRccCab(ByVal psServer As String, ByVal sFecCD As String)
Dim sql As String
Dim Co As DConecta

Set Co = New DConecta

sql = " Insert " & psServer & "rcctotal"

sql = sql & " Select"
sql = sql & " convert(datetime,substring(Col001,12,8)) Fec_Rep,"
sql = sql & " substring(Col001,2,10) Cod_Edu, null Tip_Det,"
sql = sql & " substring(Col001,20,1) Tip_Doc_Trib,"
sql = sql & " substring(Col001,21,11) Cod_Doc_Trib,"
sql = sql & " substring(Col001,32,1) Tip_Doc_Id,"
sql = sql & " substring(Col001,33,12) Cod_Doc_Id,"
sql = sql & " substring(Col001,45,1) TipPers,"
sql = sql & " substring(Col001,46,1) TipEmp,"
sql = sql & " convert(int,substring(Col001,47,3)) Can_ents,"
sql = sql & " isnull(convert(real,substring(Col001,50,5)/100),'') Calif_0,"
sql = sql & " isnull(convert(real,substring(Col001,55,5)/100),'') Calif_1,"
sql = sql & " isnull(convert(real,substring(Col001,60,5)/100),'') Calif_2,"
sql = sql & " isnull(convert(real,substring(Col001,65,5)/100),'') Calif_3,"
sql = sql & " isnull(convert(real,substring(Col001,70,5)/100),'') Calif_4,"
sql = sql & " case when len(rtrim(ltrim(substring(Col001,75,120))))=0 then '' else rtrim(ltrim(substring(Col001,75,120))) + ' ' end +"
sql = sql & " case when len(IsNull(rtrim(ltrim(substring(Col001,195,40))),'')) = 0 then '' else  + IsNull(rtrim(ltrim(substring(Col001,195,40))),'')  + ' ' end +"
sql = sql & " case when len(rtrim(ltrim(substring(Col001,235,40))))=0 then '' else rtrim(ltrim(substring(Col001,235,40))) + ' ' end +"
sql = sql & " case when len(rtrim(ltrim(substring(Col001,275,40))))=0 then '' else rtrim(ltrim(substring(Col001,275,40))) + ' ' end +"
sql = sql & " case when len(rtrim(ltrim(substring(Col001,315,40))))=0 then '' else rtrim(ltrim(substring(Col001,315,40))) end Nom_Deu"
sql = sql & " From " & psServer & "rcc" & sFecCD
sql = sql & " where substring(Col001,1,1)='1'"


On Error GoTo RccErr
    Co.AbreConexion
    'Co.AdicionaCmdBatch (Sql)
    Co.ConexionActiva.CommandTimeout = 12000
    Co.Ejecutar (sql)
    
    Co.CierraConexion
    Set Co = Nothing
Exit Sub

RccErr:
   Call RaiseError(MyUnhandledError, "DRcc:InsertaRccDet Method")
   Set Co = Nothing
End Sub

Sub InsertRCCDet(ByVal psServer As String, ByVal sFecCD As String)
Dim sql As String
Dim sF As String
Dim Co As DConecta

Set Co = New DConecta
'Calcula ultimo dia del mes
sF = "01/" & Right(sFecCD, 2) & "/" & Left(sFecCD, 4)
sF = DateAdd("M", 1, sF)
sF = DateAdd("D", -1, sF)

sql = "Insert " & psServer & "rcctotaldet"
sql = sql & "(dFecha,Cod_Edu,Cod_Emp,Tip_Credito, Cod_Cuenta,Condicion,Val_Saldo,Clasificacion)"

'Sql = Sql & " Select   convert(datetime,'" & Format(sF, "YYYY/MM/DD") & "'),substring(Col001,3,10)codedu,"
'Sql = Sql & " isnull(substring(Col001,2,1),'')tipdet,isnull(substring(Col001,13,5),'')codemp,"
'Sql = Sql & " isnull(substring(Col001,18,1),'')tipcredito,"
'Sql = Sql & " isnull(substring(Col001,19,14),'')codcuenta,"
'Sql = Sql & " case when isnumeric(substring(Col001,33,4)) = 1 then substring(Col001,33,4) else  null end condicion,"
'Sql = Sql & " convert(money,substring(Col001,37,15))/100 valsaldo,"
'Sql = Sql & " isnull(substring(Col001,52,1),'')clasificacion"
'Sql = Sql & " From " & psServer & "rcc" & sFecCD
'Sql = Sql & " where substring(Col001,1,1)='2'"


sql = sql & " Select"
sql = sql & " convert(datetime,'" & Format(sF, "YYYY/MM/DD") & "') dFecha,"
sql = sql & " substring(Col001,2,10) Cod_Edu,"
sql = sql & " substring(Col001,12,5) Cod_Emp,"
sql = sql & " substring(Col001,17,1) Tip_Credito,"
sql = sql & " isnull(substring(Col001,18,14),'') Cod_Cuenta,"
sql = sql & " case when isnumeric(substring(Col001,32,4)) = 1 then substring(Col001,32,4) else  null end Condicion,"
sql = sql & " convert(money,substring(Col001,36,15))/100 Val_Saldo,"
sql = sql & " substring(Col001,51,1) Clasificacion"
sql = sql & " From " & psServer & "rcc" & sFecCD
sql = sql & " where substring(Col001,1,1)='2'"


On Error GoTo RccErr
    Co.AbreConexion
    Co.ConexionActiva.CommandTimeout = 12000
    Co.Ejecutar (sql)
    Co.CierraConexion
    Set Co = Nothing
Exit Sub

RccErr:
   Call RaiseError(MyUnhandledError, "DRcc:InsertaRccDet Method")
   Set Co = Nothing
End Sub

Sub InsertaDataColocCalifProvTotal(ByVal sServer As String, ByVal psFecData As String)
Dim sql As String
Dim Co As DConecta
Set Co = New DConecta
sql = " INSERT " & sServer & "ColocCalifProvTotal "
sql = sql & " SELECT '" & Format(psFecData, "YYYY/MM/DD") & "', * FROM ColocCalifProv"
Co.AbreConexion
Co.Ejecutar (sql)
Co.CierraConexion
Set Co = Nothing
End Sub

Sub ActulizaFechaCierre(ByVal psFecData As String)
Dim sql As String
Dim Co As DConecta
Set Co = New DConecta
Co.AbreConexion
sql = "UPDATE CONSTSISTEMA SET NCONSSISVALOR= '" & psFecData & "' where nConsSisCod = 160"
Co.Ejecutar (sql)
Co.CierraConexion
Set Co = Nothing
End Sub
Function VerificaDataMigradaFecha(ByVal psServer As String, ByVal psFecData As String) As Long
Dim sql As String
Dim Co As DConecta
Dim rs As ADODB.Recordset

Set Co = New DConecta

sql = " select count(*) nro from " & psServer & "ColocCalifProvTotal "
sql = sql & " where datediff(day,dfecha,'" & Format(psFecData, "YYYY/MM/DD") & "')=0"

Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
VerificaDataMigradaFecha = rs!Nro
If rs!Nro > 0 Then
    MsgBox "La Data ya fue Transferida", vbInformation, "AVISO"
    Exit Function
End If
Set Co = Nothing
Set rs = Nothing
End Function
Function FechasRcc(ByVal psServer As String) As ADODB.Recordset
Dim sql As String
Dim Co As DConecta
Dim rs As ADODB.Recordset

Set Co = New DConecta

sql = "select distinct fec_Rep Fecha from " & psServer & "rcchistotal"
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
Co.CierraConexion
Set FechasRcc = rs

Set Co = Nothing
Set rs = Nothing
End Function

Function ServConsol(ByVal pnServer As Integer) As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim Co As DConecta
sql = "Select nConsSisValor serv from ConstSistema where nConsSisCod = " & pnServer
Set Co = New DConecta
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
If Not (rs.EOF And rs.BOF) Then
    ServConsol = Trim(rs!Serv)
End If
Co.CierraConexion
Set rs = Nothing
Set Co = Nothing
End Function

Function NombreServerConsol() As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta
sql = "SELECT Substring(nConsSisValor,1,LEN(nConsSisValor) - 5) SERVIDOR FROM ConstSistema WHERE nConsSisCod  = 43"
Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)
If Not (rs.EOF And rs.BOF) Then
    NombreServerConsol = Trim(rs!SERVIDOR)
End If
Co.CierraConexion
Set rs = Nothing
Set Co = Nothing
End Function

Function VerificaTabla(ByVal psServer As String, ByVal sFecCD As String) As Integer
Dim sql As String
Dim Co As DConecta
Dim rs As ADODB.Recordset

sql = " select * from " & psServer & "sysobjects  where name = 'Rcc" & sFecCD & "'"
Set Co = New DConecta
Co.AbreConexion


Set rs = Co.CargaRecordSet(sql)
If Not (rs.EOF And rs.BOF) Then
    VerificaTabla = 1
Else
    VerificaTabla = -1
End If
Co.CierraConexion
Set Co = Nothing
Set rs = Nothing
End Function


