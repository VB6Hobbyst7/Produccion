VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMGargaMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit




'Documentos
Public gnDocTpo  As Long
Public gsDocDesc As String
Public gsDocNro  As String

'* Para el mantenimiento de Permisos


Public Function LimpiaVarLavDinero()
    gReaPersLavDinero = ""
    gBenPersLavDinero = ""
End Function

'' PinPad
''Public myPSerial                As HCOMPINPADLib.Pinpad
Private Function DamePosicionNivel(ByVal psName As String) As Integer
Dim i As Integer
Dim Y As Integer
    Y = 1
    For i = 4 To Len(psName) Step 2
        If Mid(psName, i, 2) = "00" Then
            DamePosicionNivel = Y
            Exit For
        End If
        Y = Y + 1
    Next i
    If Y = 6 Then
        DamePosicionNivel = 5
    End If
    
End Function

Public Sub CargaMenuArbol(ByRef pR As ADODB.Recordset, ByRef nPunt As Integer, ByRef pnId As Integer)
Dim nPos As Integer
Dim nPos2 As Integer
Dim nPuntTemp As Integer

        If pR.EOF Then
            Exit Sub
        End If
        
        'Obtengo el Nivel
        nPos = DamePosicionNivel(MatMenuItems(nPunt - 1).sName)
        nPos2 = DamePosicionNivel(pR!cname)
        
        If nPos2 > nPos Then 'Es Hijo
            pnId = pnId + 1
            ReDim Preserve MatMenuItems(nPunt + 1)
            MatMenuItems(nPunt).nId = pnId - 1
            MatMenuItems(nPunt).sCodigo = Trim(pR!cCodigo)
            MatMenuItems(nPunt).sCaption = Trim(pR!cDescrip)
            MatMenuItems(nPunt).sName = Trim(pR!cname)
            MatMenuItems(nPunt).sIndex = Right(pR!cname, 2)
            MatMenuItems(nPunt).bCheck = False
            MatMenuItems(nPunt).nPuntDer = -1
            MatMenuItems(nPunt).nPuntAbajo = -1
            MatMenuItems(nPunt - 1).nPuntDer = MatMenuItems(nPunt).nId
            MatMenuItems(nPunt).nNivel = nPos2
            pR.MoveNext
            
            nPunt = nPunt + 1
            Call CargaMenuArbol(pR, nPunt, pnId)
            
        End If
        
        If nPos2 = nPos Then 'Son del Mismo Menu
            pnId = pnId + 1
            nPuntTemp = nPunt
            If pnId <> nPunt + 1 Then
                nPunt = pnId - 1
            End If
            
            ReDim Preserve MatMenuItems(nPunt + 1)
            MatMenuItems(nPunt).nId = pnId - 1
            MatMenuItems(nPunt).sCodigo = Trim(pR!cCodigo)
            MatMenuItems(nPunt).sCaption = Trim(pR!cDescrip)
            MatMenuItems(nPunt).sName = Trim(pR!cname)
            MatMenuItems(nPunt).sIndex = Right(pR!cname, 2)
            MatMenuItems(nPunt).bCheck = False
            MatMenuItems(nPunt).nPuntDer = -1
            MatMenuItems(nPunt).nPuntAbajo = -1
            MatMenuItems(nPuntTemp - 1).nPuntAbajo = MatMenuItems(nPunt).nId
            MatMenuItems(nPunt).nNivel = nPos2
            pR.MoveNext
            
            nPunt = nPunt + 1
            Call CargaMenuArbol(pR, nPunt, pnId)
        End If
        
        If nPos2 < nPos Then 'Es un Menu de nivel Anterior
            nPunt = nPunt - 1
            Call CargaMenuArbol(pR, nPunt, pnId)
        End If
        
End Sub


Public Sub CargaVarSistema(ByVal pbContabilidad As Boolean)
    Dim lsQrySis As String
    Dim rsQrySis As New ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    Dim VSQL As String
    Dim lnStrConn As String
    Dim lnPosIni As Integer
    Dim lnPosFin As Integer
    Dim lnStr As String
    Set oConect = New COMConecta.DCOMConecta
    
    If oConect.AbreConexion(gsConnection) = False Then
        Exit Sub
    End If
    
    lsQrySis = " Select * From ConstSistema " _
            & " Where nConsSisCod in (" & gConstSistFechaSistema & "," & gConstSistNombreAbrevCMAC & "," _
            & gConstSistRutaBackup & "," & gConstSistCodCMAC & "," & gConstSistMargenSupCartas & "," & gConstSistMagenIzqCartas & "," _
            & gConstSistMargenDerCartas & "," & gConstSistNroLineasPagina & "," & gConstSistNroLineasOrdenPago & "," _
            & gConstSistCtaConversionMEDol & "," & gConstSistCtaConversiónMESoles & "," & gConstSistTipoConverión & "," _
            & gConstSistNombreModulo & "," & gConstSistFechaInicioDia & "," & gConstSistDominio & "," & gConstSistPDC _
            & ",40,50," & gConstPersCodCMACT & "," & gConstSistAgenciaEspecial & "," & gConstSistCierreMesNegocio & ", " _
            & gConstSistRutaIcono & ",84,85," & gConstSistVerificaRegistroEfectivo & ") " _
            & "ORDER BY nConsSisCod"
    
    Set rsQrySis = oConect.CargaRecordSet(lsQrySis)
    If rsQrySis.BOF Or rsQrySis.EOF Then
       rsQrySis.Close
       Set rsQrySis = Nothing
       MsgBox "Tabla VarSistema está vacia", vbInformation, "Aviso"
       gdFecSis = ""
       gsInstCmac = ""
       gsNomCmac = ""
       gsCodCMAC = ""
       Exit Sub
    End If
    Do While Not rsQrySis.EOF
        Select Case Trim(rsQrySis!nConsSisCod)
                Case gConstSistFechaSistema
                        gdFecSis = CDate(Trim(rsQrySis!nConsSisValor))
                Case gConstSistCierreMesNegocio
                        gdFecDataFM = CDate(Trim(rsQrySis!nConsSisValor))
                        gdFecData = CDate(Trim(rsQrySis!nConsSisValor))
                Case gConstSistNombreAbrevCMAC
                        gsInstCmac = Trim(rsQrySis!nConsSisValor)
                        gsNomCmac = Trim(rsQrySis!nConsSisDesc)
                Case gConstSistCodCMAC
                        gsCodCMAC = Trim(rsQrySis!nConsSisValor)
                Case gConstSistRutaBackup
                        gsDirBackup = Trim(rsQrySis!nConsSisValor)
                Case gConstSistNombreModulo '   "cEmpresa":
                        gcTitModulo = Trim(rsQrySis!nConsSisDesc)
                        gcModuloLogo = Trim(rsQrySis!nConsSisValor)
                Case gConstSistMargenSupCartas  ' "nMargSup":
                        gnMgSup = Val(rsQrySis!nConsSisValor)
                Case gConstSistMagenIzqCartas  ' "nMargIzq":
                        gnMgIzq = Val(rsQrySis!nConsSisValor)
                Case gConstSistMargenDerCartas  '  "nMargDer":
                        gnMgDer = Val(rsQrySis!nConsSisValor)
                Case gConstSistNroLineasPagina  '      "nLinPage":
                        gnLinPage = Val(rsQrySis!nConsSisValor)
                Case gConstSistNroLineasOrdenPago  '   "nLinPageOP":
                        gnLinPageOP = Val(rsQrySis!nConsSisValor)
                Case gConstSistCtaConversionMEDol '   "cConvMED":
                        gcConvMED = Trim(rsQrySis!nConsSisValor)
                Case gConstSistCtaConversiónMESoles '    "cConvMES":
                        gcConvMES = Trim(rsQrySis!nConsSisValor)
                Case gConstSistTipoConverión  '  "cConvTipo":
                        gcConvTipo = Trim(rsQrySis!nConsSisValor)
                Case gConstSistDominio
                        gsDominio = Trim(rsQrySis!nConsSisValor)
                Case gConstSistPDC
                        gsPDC = Trim(rsQrySis!nConsSisValor)
                Case gConstPersCodCMACT
                    gsCodPersCMACT = Trim(rsQrySis!nConsSisValor)
                Case 40
                    gcCtaIGV = Trim(rsQrySis!nConsSisValor)
                Case 50
                    gcEmpresaRUC = Trim(rsQrySis!nConsSisValor)
                Case gConstSistAgenciaEspecial
                    gbAgeEsp = IIf(Trim(rsQrySis!nConsSisValor) = "1", True, False)
                Case gConstSistRutaIcono
                    gsRutaIcono = Trim(rsQrySis!nConsSisValor)
                Case gConstSistVerificaRegistroEfectivo
                    gbVerificaRegistroEfectivo = IIf(rsQrySis!nConsSisValor = "1", True, False)
                Case 84
                  gnValidaSolCredito = rsQrySis!nConsSisValor
                
                Case 85
                    gnValidaGarantia = rsQrySis!nConsSisValor
                    
        End Select
        rsQrySis.MoveNext
    Loop
    rsQrySis.Close
    Set rsQrySis = Nothing
    
    'Deduce el nombre del Servidor
    
    gsServerName = oConect.servername
    'Deduce el nombre de la Base de Datos
    gsDBName = oConect.DatabaseName
    lnStrConn = oConect.CadenaConexion
    'Deduce el nombre de usuario
    lnPosIni = InStr(1, lnStrConn, "UID=", vbTextCompare)
    If lnPosIni > 0 Then
        lnPosFin = InStr(lnPosIni, lnStrConn, ";", vbTextCompare)
        lnStr = Mid(lnStrConn, lnPosIni, lnPosFin - lnPosIni)
        lnPosIni = InStr(1, lnStr, "=", vbTextCompare)
        gsUID = Mid(lnStr, lnPosIni + 1, Len(lnStr) - lnPosIni)
    Else
        gsUID = ""
    End If
    'Deduce el password
    lnPosIni = InStr(1, lnStrConn, "PWD=", vbTextCompare)
    If lnPosIni > 0 Then
        lnPosFin = InStr(lnPosIni, lnStrConn, ";", vbTextCompare)
        lnStr = Mid(lnStrConn, lnPosIni, lnPosFin - lnPosIni)
        lnPosIni = InStr(1, lnStr, "=", vbTextCompare)
        gsPWD = Mid(lnStr, lnPosIni + 1, Len(lnStr) - lnPosIni)
    Else
        gsPWD = ""
    End If
    oConect.CierraConexion
    Set oConect = Nothing
    
    gcMN = "S/."
    gcME = "$"
    
    gcMNDig = "1"
    gcMEDig = "2"
End Sub



Public Sub OpeDiaCreaTemporal(psFecha As String, psCodUsu As String)
Dim sql As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

oCon.AbreConexion

sql = " if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[##Mov_" & psCodUsu & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" _
& " drop table [dbo].[##Mov_" & psCodUsu & "]" _
& " "
oCon.Ejecutar sql

sql = " CREATE TABLE [dbo].[##Mov_" & psCodUsu & "] (" _
& " [nMovNro] [int] NOT NULL ," _
& " [cMovNro] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," _
& " [cOpeCod] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," _
& " [cMovDesc] [varchar] (302) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," _
& " [nMovEstado] [int] NOT NULL ," _
& " [nMovFlag] [int] NULL" _
& " ) ON [PRIMARY]" _
& " "
oCon.Ejecutar sql
sql = " ALTER TABLE [dbo].[##Mov_" & psCodUsu & "] WITH NOCHECK ADD" _
& " CONSTRAINT [PK_Mov_" & psCodUsu & "] PRIMARY KEY CLUSTERED" _
& " (" _
& " [nMovNro]" _
& " ) WITH FILLFACTOR = 90 ON [PRIMARY]" _
& " "
oCon.Ejecutar sql

sql = " Insert Into [dbo].[##Mov_" & psCodUsu & "] (nMovNro, cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag)" _
& " Select nMovNro, cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag From Mov Where cMovNro Like '" & psFecha & "%'"
oCon.Ejecutar sql

oCon.CierraConexion

End Sub
Public Sub OpeDiaEliminaTemporal(psCodUsu As String)
Dim sql As String
Dim oCon As COMConecta.DCOMConecta
Set oCon = New COMConecta.DCOMConecta

oCon.AbreConexion

sql = " drop table [dbo].[##Mov_" & psCodUsu & "]"
oCon.Ejecutar sql

oCon.CierraConexion
End Sub




