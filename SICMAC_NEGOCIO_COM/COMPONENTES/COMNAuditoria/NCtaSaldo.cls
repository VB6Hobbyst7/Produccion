VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCtasaldo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A82E291008C"
Option Base 0
Option Explicit
Dim dbConec As COMConecta.DCOMConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sCentralCom As String

Public Function GetCtaSaldo(psCtaContCod As String, psFecha As String, Optional pbMN As Boolean = True) As Currency
   On Error GoTo GetCtaSaldoErr
   Dim lsCond As String
   If Right(psCtaContCod, 1) = "%" Then
        lsCond = " LIKE "
   Else
        lsCond = " = "
   End If
   
   If lsCond <> " = " Then
        If Mid(psCtaContCod, 3, 1) = "0" Then
            psSql = "Select dbo.GetSaldoCtaAcumulado('" & psFecha & "','" & Left(psCtaContCod, 2) & "_" & Mid(psCtaContCod, 4) & "'," & IIf(pbMN, "1", "2") & ") as Saldo"
            Set prs = dbConec.CargaRecordSet(psSql)
            If Not prs.EOF Then
               If pbMN Then '
                  GetCtaSaldo = prs!Saldo
               Else
                  GetCtaSaldo = prs!Saldo
               End If
            Else
               GetCtaSaldo = 0
            End If
            
        Else
            psSql = "SELECT ISNULL(SUM(nCtaSaldoImporte),0) nCtaSaldoImporte, ISNULL(SUM(nCtaSaldoImporteME),0) nCtaSaldoImporteME " _
                  & "FROM CtaSaldo cs WHERE cCtaContCod " & lsCond & "'" & psCtaContCod & "' AND dCtaSaldoFecha = " _
                  & "      (SELECT MAX(dCtaSaldoFecha) FROM CtaSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and dCtaSaldoFecha <= '" & psFecha & "')"
        
             Set prs = dbConec.CargaRecordSet(psSql)
             If Not prs.EOF Then
                If pbMN Then
                   GetCtaSaldo = prs!nCtaSaldoImporte
                Else
                   GetCtaSaldo = prs!nCtasaldoImporteME
                End If
             Else
                GetCtaSaldo = 0
             End If
        End If
    Else
       If pbMN Then
          psSql = "SELECT dbo.getsaldocta('" & psFecha & "','" & psCtaContCod & "',1) "
       Else
          psSql = "SELECT dbo.getsaldocta('" & psFecha & "','" & psCtaContCod & "',2) "
       End If
    
       Set prs = dbConec.CargaRecordSet(psSql)
       GetCtaSaldo = prs.Fields(0)
       
    End If
   Exit Function
GetCtaSaldoErr:
   MsgBox Err.Description, vbCritical
End Function

Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As New COMConecta.DCOMClasIni
   sCentralCom = oIni.BaseComunes
Set oIni = Nothing

Set dbConec = New COMConecta.DCOMConecta
dbConec.AbreConexion
End Sub

Private Sub Class_Terminate()
dbConec.CierraConexion
Set dbConec = Nothing
End Sub
