VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCaja_Adeudados"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A89DAEE0013"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim vsConexion As String
Dim vsServerAdmin As String
Dim vsServerComunes As String
Dim vsServerImagenes As Variant
Dim vsServerNegocio As String
Dim vsServerPersonas As String
    
Dim sql As String
Dim oConect As COMConecta.DCOMConecta
Dim rs As ADODB.Recordset

Public Function CargaIndiceVAC(pdFecha As Date) As Double
'    On Error GoTo CargaIndiceVACErr
'    CargaIndiceVAC = 0
'    sql = "Select * from IndiceVac iv Where dIndiceVac = (Select Max(dIndiceVac) FROM IndiceVac iv1 where iv1.dIndiceVac <= '" & Format(pdFecha, gsFormatoFecha) & "') "
'    Set rs = oConect.CargaRecordSet(sql)
'    If Not rs.EOF And Not rs.BOF Then
'        CargaIndiceVAC = Format(rs!nIndiceVac, "#,###.00####")
'    End If
'    RSClose rs
'
'    Exit Function
'CargaIndiceVACErr:
'    Call RaiseError(MyUnhandledError, "DCaja_Adeudados:CargaIndiceVAC Method")

    On Error GoTo CargaIndiceVACErr
    CargaIndiceVAC = 0
    sql = "Select dIndiceVac, ISNULL(nIndiceVac,0) as nIndiceVac from IndiceVac iv Where dIndiceVac = (Select Max(dIndiceVac) FROM IndiceVac iv1 where iv1.dIndiceVac = '" & Format(pdFecha, gsFormatoFecha) & "') "
    Set rs = oConect.CargaRecordSet(sql)
    If Not rs.EOF And Not rs.BOF Then
        CargaIndiceVAC = Format(rs!nIndiceVac, "#,###.00####")
    End If
    RSClose rs
    
    Exit Function
CargaIndiceVACErr:
    'Call RaiseError(MyUnhandledError, "DCaja_Adeudados:CargaIndiceVAC Method")
    MsgBox Err.Description, vbCritical
End Function



Public Function GetCuotaMalPagador(psPersCod As String, psIFTpo As String, psCtaIfCod As String) As ADODB.Recordset
    Dim sSQL As String
    Dim oCon As New COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    
    sSQL = "Select cperscod,ciftpo,cctaifcod,nCapital,nInteres,ncomision,nnrocuota from ctaifcalendario"
    sSQL = sSQL & " where cperscod='" & psPersCod & "' and ciftpo='" & psIFTpo & "' and cctaifcod='" & psCtaIfCod & "'  and nnrocuota in (  select min(nnrocuota)"
    sSQL = sSQL & "     From ctaifcalendario"
    sSQL = sSQL & "     where cperscod='" & psPersCod & "' and ciftpo='" & psIFTpo & "' and cctaifcod='" & psCtaIfCod & "'"
    sSQL = sSQL & "     and ctpocuota=6 and cestado in (0,2)"
    sSQL = sSQL & "           ) and ctpocuota=6"
'    ocon.AbreConexion
    Set rs = oConect.CargaRecordSet(sSQL)
    Set GetCuotaMalPagador = rs
'    ocon.CierraConexion
End Function


Public Function CalculaInteres(pnDias As Long, pnPeriodo As Long, pnTasaInt As Currency, pnCapital As Currency) As Currency
    Dim lnFactor As Double
    Dim lnBaseInt As Double
    If pnPeriodo > 0 Then
        lnFactor = pnDias / pnPeriodo
        lnBaseInt = ((1 + (pnTasaInt / 100)) ^ lnFactor) - 1
        CalculaInteres = lnBaseInt * pnCapital
    Else
        CalculaInteres = 0
    End If
End Function


Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As COMConecta.DCOMClasIni

Set oIni = New COMConecta.DCOMClasIni
vsServerAdmin = oIni.BaseAdministracion
vsServerComunes = oIni.BaseComunes
vsServerImagenes = oIni.BaseImagenes
vsServerNegocio = oIni.BaseNegocio
vsServerPersonas = oIni.BasePersonas
    
Set oConect = New COMConecta.DCOMConecta
oConect.AbreConexion

End Sub

Private Sub Class_Terminate()
oConect.CierraConexion
Set oConect = Nothing
End Sub


