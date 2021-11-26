VERSION 5.00
Begin VB.Form frmColPDesemCampAdjudicado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desembolso Campaña adjudicado"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "frmColPDesemCampAdjudicado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5130
      TabIndex        =   15
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6400
      TabIndex        =   13
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   4340
      Width           =   7510
      Begin VB.Frame Frame3 
         Caption         =   "Nº Contrato Adjudicado"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2775
         Begin VB.Label lblCtaCMAC 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   23
            Top             =   270
            Width           =   400
         End
         Begin VB.Label lblCtaAGE 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   675
            TabIndex        =   22
            Top             =   270
            Width           =   300
         End
         Begin VB.Label lblCtaPRD 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1005
            TabIndex        =   21
            Top             =   270
            Width           =   400
         End
         Begin VB.Label lblCtaNum 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   20
            Top             =   270
            Width           =   1170
         End
      End
      Begin VB.Label lblTotalCobrar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6000
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblITF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6000
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblMontoNewCred 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblInicial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6000
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Total Cobrar a Cliente:"
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   1000
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "ITF:"
         Height          =   255
         Left            =   4320
         TabIndex        =   11
         Top             =   650
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Inicial:"
         Height          =   255
         Left            =   4320
         TabIndex        =   10
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Monto Nuevo Crédito:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nº Contrato"
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtCtaCMAC 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   270
         Width           =   400
      End
      Begin VB.TextBox txtCtaAGE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   675
         TabIndex        =   4
         Top             =   270
         Width           =   300
      End
      Begin VB.TextBox txtCtaPRD 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1005
         TabIndex        =   3
         Top             =   270
         Width           =   400
      End
      Begin VB.TextBox txtCtaNUM 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   270
         Width           =   1170
      End
      Begin VB.CommandButton cmdBuscarCliente 
         Caption         =   "..."
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   270
         Width           =   280
      End
   End
   Begin SICMACT.ActXColPDesCon AXDesCon 
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6165
   End
End
Attribute VB_Name = "frmColPDesemCampAdjudicado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCtaCod As String
Dim sPersCod As String

'DATOS CREDITO ADJUDICADO
Dim sCredAdjCod As String
Dim nCredAdjSaldCap As Double
Dim nCredAdjnIntComp As Double
Dim nCredAdjnIntCompVenc As Double
Dim nCredAdjnIntMora As Double
Dim nCredAdjnCostoNotif As Double
Dim nCredAdjnValTasacion As Double
Dim nCredAdjnValdeudaTot As Double
Dim nCredAdj14k As Double
Dim nCredAdj16k As Double
Dim nCredAdj18k As Double
Dim nCredAdj21k As Double
'FIN DATOS

Private Sub cmdBuscarCliente_Click()
    Dim loPersContrato As New COMDColocPig.DCOMColPContrato
    Dim lrContratos As New ADODB.Recordset
    Dim loCuentas As New COMDPersona.UCOMProdPersona
    Dim lsEstados As String
    Dim loPers As COMDPersona.UCOMPersona
    Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    sPersCod = loPers.sPersCod
    lsEstados = "2100"
    If Trim(sPersCod) <> "" Then
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(sPersCod, lsEstados, Mid(gsCodAge, 4, 2))
        Set loPersContrato = Nothing
    End If
    Set loCuentas = frmProdPersona.Inicio(loPers.sPersNombre, lrContratos)
    If loCuentas.sCtaCod <> "" Then
        txtCtaCMAC = Mid(loCuentas.sCtaCod, 1, 3)
        txtCtaAGE = Mid(loCuentas.sCtaCod, 4, 2)
        txtCtaPRD = Mid(loCuentas.sCtaCod, 6, 3)
        txtCtaNUM = Mid(loCuentas.sCtaCod, 9, 18)
        sCtaCod = loCuentas.sCtaCod
    End If
    Set loCuentas = Nothing
    Set loPers = Nothing
End Sub

Private Sub BuscaContrato(ByVal psNroContrato As String)
    Dim loValContrato As New COMNColoCPig.NCOMColPValida
    Dim ClsPersona As New COMDPersona.DCOMPersonas
    Dim oCreditoN As New COMNCredito.NCOMCredito
    Dim oPers  As New COMDPersona.UCOMPersona
    Dim oDCredito As New COMDCredito.DCOMCredito
    Dim oColP As New COMNColoCPig.NCOMColPContrato
    Dim lrValida As New ADODB.Recordset
    Dim lrsDatosDesemb As New ADODB.Recordset
    Dim Rf As New ADODB.Recordset
    Dim loValMontoPrestamo As Double
    Dim loValOtrosCostos As Double
    Dim lsmensaje As String
    Dim sPersCodTitular As String
    Dim lbok As Boolean
    Dim lcCodPersTitu As String
    Dim loItf1 As Double
    Dim nRedondeoITF As Double
    Dim loItf2 As Double
    Dim lcCodCli As String
    sPersCodTitular = oDCredito.RecuperaTitularCredito(psNroContrato)
    Set lrsDatosDesemb = oColP.DevuelveValorDesemCampAdjudicado(psNroContrato)
    If oPers.fgVerificaEmpleado(sPersCodTitular) Or oPers.fgVerificaEmpleadoVincualdo(sPersCodTitular) Then
        If Not oCreditoN.ExisteAsignaSaldo(psNroContrato, 2) Then
            MsgBox "El crédito aún no tiene saldo asignado, verificar con el Departamente de Administración de Créditos.", vbInformation, "Aviso"
            Exit Sub
        End If
        Set oCreditoN = Nothing
    End If
    Set oPers = Nothing

    gITF.fgITFParamAsume (Mid(psNroContrato, 4, 2)), Mid(psNroContrato, 6, 3)
    Set lrValida = loValContrato.nValidaDesembolsoCredPignoraticio(psNroContrato, gdFecSis, 0, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    Set loValContrato = Nothing
    lcCodPersTitu = lrValida!cPersCod
    If lrValida Is Nothing Then
        'Limpiar
        Set lrValida = Nothing
        Exit Sub
    End If
    
    lbok = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False)
    Set lrValida = Nothing
    loValMontoPrestamo = CCur(AXDesCon.SaldoCapital)
    
    If Not (lrsDatosDesemb.EOF And lrsDatosDesemb.BOF) Then
         lblInicial.Caption = lrsDatosDesemb!nMontoInicial
         lblCtaCMAC.Caption = Mid(lrsDatosDesemb!cCtaCodAdj, 1, 3)
         lblCtaAGE.Caption = Mid(lrsDatosDesemb!cCtaCodAdj, 4, 2)
         lblCtaPRD.Caption = Mid(lrsDatosDesemb!cCtaCodAdj, 6, 3)
         lblCtaNum.Caption = Mid(lrsDatosDesemb!cCtaCodAdj, 9, 10)
         lblMontoNewCred.Caption = Me.AXDesCon.SaldoCapital
         sCredAdjCod = lrsDatosDesemb!cCtaCodAdj
    End If
    Call CargarDatosCreditoAdjudicado(sCredAdjCod)
    If gITF.gbITFAplica Then
        If Not gITF.gbITFAsumidocreditos Then
            loItf1 = Format(gITF.fgITFCalculaImpuesto(Val(Me.lblInicial.Caption)), "#0.00")
            Me.LblItf.Caption = Format(loItf1, "#0.00")
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblItf.Caption))
            If nRedondeoITF > 0 Then
               Me.LblItf.Caption = Format(CCur(Me.LblItf.Caption) - nRedondeoITF, "#,##0.00")
               loItf1 = LblItf.Caption
            End If
            Me.lblTotalCobrar.Caption = Format(CDbl(Me.lblInicial.Caption) + CDbl(Me.LblItf.Caption), "#0.00")
        Else
            loItf1 = Format(gITF.fgITFCalculaImpuesto(loValMontoPrestamo), "#0.00")
            loItf2 = Format(gITF.fgITFCalculaImpuesto(loValOtrosCostos), "#0.00")
            Me.LblItf = Format(loItf1 + loItf2, "#0.00")
            nRedondeoITF = fgDiferenciaRedondeoITF(CCur(Me.LblItf.Caption))
            If nRedondeoITF > 0 Then
               Me.LblItf = Format(CCur(Me.LblItf.Caption) - nRedondeoITF, "#,##0.00")
               If loItf1 >= nRedondeoITF Then
                    loItf1 = loItf1 - nRedondeoITF
               Else
                    loItf2 = loItf2 - nRedondeoITF
               End If
            End If

            Me.lblTotalCobrar.Caption = Me.lblInicial.Caption
        End If
    Else
        Me.LblItf.Caption = Format(0, "#0.00")
        Me.lblTotalCobrar.Caption = Me.lblInicial.Caption
    End If
    

    
    Set Rf = ClsPersona.BuscaCliente(lcCodCli, BusquedaCodigo)
    If Not Rf.BOF And Not Rf.EOF Then
        If Rf!nPersPersoneria = 1 Then
            Call frmPersonaFirma.Inicio(Trim(lcCodCli), Mid(lcCodCli, 4, 2), False, True)
        End If
    End If
    Set Rf = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiarFormulario
End Sub

Private Sub cmdGrabar_Click()
    Dim loContFunct As New COMNContabilidad.NCOMContFunciones
    Dim loGrabarDesem As New COMNColoCPig.NCOMColPContrato
    Dim loColImp As New COMNColoCPig.NCOMColPImpre
    Dim clsprevio As New previo.clsprevio
    Dim oMov As New COMDMov.DCOMMov
    Dim opt As Integer
    Dim OptBt2 As Integer
    Dim nFicSal As Integer
    Dim lnMovNro As Long
    Dim lnSaldoCap As Currency, lnInteresComp As Currency, lnImpuesto As Currency
    Dim lnCostoTasacion As Currency, lnCostoCustodia As Currency
    Dim lnMontoEntregar As Currency
    Dim lcTextImp As String
    Dim lsMovNro As String
    Dim lsFechaHoraGrab As String
    Dim lsFechaHoraPrend As String
    Dim lsCuenta As String
    Dim lsCadImp As String
    Dim lcNomCli As String
    Dim sOperacion As String
    Dim nRedondeoITF As Double
    lsCuenta = sCtaCod
    lnSaldoCap = Me.AXDesCon.SaldoCapital
    lnInteresComp = 0
    lnImpuesto = 0
    lnCostoCustodia = 0
    lnCostoTasacion = 0
    lnMontoEntregar = CCur(Me.lblTotalCobrar.Caption)
        
    Dim objValidaLimiteTpProducto As COMDColocPig.DCOMColPContrato 'JOEP ERS047 20170904
    Dim RsTpProducto As ADODB.Recordset 'JOEP ERS047 20170904
    Dim RsZonaGeog As ADODB.Recordset 'JOEP ERS047 20170904
    
    Dim oPersona As New COMNPersona.NCOMPersona
    If oPersona.NecesitaActualizarDatos(sPersCod, gdFecSis) Then
         MsgBox "Para continuar con la Operación Ud. debe actualizar los datos del" & Chr(13) & "Titular: " & lcNomCli, vbInformation, "Aviso"
         Dim foPersona As New frmPersona
         If Not foPersona.realizarMantenimiento(sPersCod) Then
             MsgBox "No se ha realizado la actualización de los datos de " & lcNomCli & "," & Chr(13) & "la Operación no puede continuar!", vbInformation, "Aviso"
             Exit Sub
         End If
    End If
    
    Dim oDPersona As New COMDPersona.DCOMPersona
    Dim rsPersonaCred As New ADODB.Recordset
    Dim rsPersona As New ADODB.Recordset
    Dim Cont As Integer
    
    Set rsPersonaCred = oDPersona.ObtenerPersCuentaRelac(sCtaCod, gColRelPersTitular)
    
    If rsPersonaCred.RecordCount > 0 Then
        If Not (rsPersonaCred.EOF And rsPersonaCred.BOF) Then
            For Cont = 0 To rsPersonaCred.RecordCount - 1
                Set rsPersona = oDPersona.ObtenerUltimaVisita(Trim(rsPersonaCred!cPersCod))
                If rsPersona.RecordCount > 0 Then
                    If Not (rsPersona.EOF And rsPersona.BOF) Then
                        If Trim(rsPersona!sUsual) = "3" Then
                        MsgBox PstaNombre(Trim(rsPersonaCred!cPersNombre), True) & "." & Chr(10) & "CLIENTE OBSERVADO: " & Trim(rsPersona!cVisObserva), vbInformation, "Aviso"
                            Call frmPersona.Inicio(Trim(rsPersonaCred!cPersCod), PersonaActualiza)
                        End If
                    End If
                End If
                Set rsPersona = Nothing
                rsPersonaCred.MoveNext
            Next Cont
        End If
    End If
    
'JOEP ERS047 20170904
Set objValidaLimiteTpProducto = New COMDColocPig.DCOMColPContrato
Set RsZonaGeog = objValidaLimiteTpProducto.ValidaLimiteZonaGeog(sCtaCod, Mid(sCtaCod, 6, 3))
Set objValidaLimiteTpProducto = Nothing
        If Not (RsZonaGeog.EOF And RsZonaGeog.BOF) Then
            If RsZonaGeog!nEstado = 0 Then
                MsgBox "El crédito supera el porcentaje máximo por Zona Geografica. Cualquier consulta comunicarse con Riesgos.", vbInformation, "Aviso"
                Exit Sub
            ElseIf RsZonaGeog!nEstado = 2 Then
                MsgBox "El crédito no puede ser sugerido, pues su solicitud de autorización fue rechazada por la gerencia de riesgos", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
RsZonaGeog.Close
'Producto
Set objValidaLimiteTpProducto = New COMDColocPig.DCOMColPContrato
Set RsTpProducto = objValidaLimiteTpProducto.ValidaLimiteTpProducto(sCtaCod, Mid(sCtaCod, 6, 3))
Set objValidaLimiteTpProducto = Nothing
    If Not (RsTpProducto.EOF And RsTpProducto.BOF) Then
        If RsTpProducto!nEstado = 0 Then
            MsgBox "El crédito supera el porcentaje máximo por Tipo de Producto Pignoraticio. Cualquier consulta comunicarse con Riesgos.", vbInformation, "Aviso"
            Exit Sub
        ElseIf RsTpProducto!nEstado = 2 Then
            MsgBox "El crédito no puede ser sugerido, pues su solicitud de autorización fue rechazada por la gerencia de riesgos", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
RsTpProducto.Close
'JOEP ERS047 20170904
    
    
    If MsgBox(" Grabar Desembolso de Contrato Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        Screen.MousePointer = 11
        cmdgrabar.Enabled = False
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
    
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        lsFechaHoraPrend = fgFechaHoraPrend(lsMovNro)
            
        Dim clsExo As New COMNCaptaServicios.NCOMCaptaServicios
        If Not clsExo.EsCuentaExoneradaLavadoDinero(lsCuenta) Then
            Dim sPersLavDinero As String
            Dim nMontoLavDinero As Double, nTC As Double
            Dim clsLav As New COMNCaptaGenerales.NCOMCaptaDefinicion, nmoneda As Integer, nMonto As Double
            nMonto = CDbl(lblTotalCobrar.Caption)
            Set clsExo = Nothing
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            nmoneda = gMonedaNacional
            If nmoneda = gMonedaNacional Then
                Dim clsTC As COMDConstSistema.NCOMTipoCambio
                Set clsTC = New COMDConstSistema.NCOMTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                sPersLavDinero = ""
            End If
        Else
            Set clsExo = Nothing
        End If
        Dim lbResultadoVisto As Boolean
        Dim loVistoElectronico As New frmVistoElectronico
        Dim bVisto As Boolean
        If loGrabarDesem.ObtieneMontoDesembCreDPigMes(sPersCod, Format(gdFecSis, "yyyyMM")) + lnSaldoCap > 5000 Then
            bVisto = True
        End If
        If bVisto = True Then
            loVistoElectronico.VistoMovNro = lsMovNro
            lbResultadoVisto = loVistoElectronico.Inicio(10, sOperacion)
            If Not lbResultadoVisto Then
                MsgBox "Operación cancelada por el usuario", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        Call loGrabarDesem.nDesembolsoCredPignoEfectivoAbono(lsCuenta, False, False, lnSaldoCap, lsFechaHoraGrab, lsMovNro, lnMontoEntregar + Val(lblMontoNewCred.Caption), 0, lnImpuesto, _
                                                lnCostoTasacion, lnCostoCustodia, IIf(False, False, gITF.gbITFAplica), gITF.gbITFAsumidocreditos, 0, 0, _
                                                CCur(lblTotalCobrar.Caption) + Val(lblMontoNewCred.Caption), lnMovNro, False)
                
        If gITF.gbITFAplica Then
            Call oMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(Me.LblItf.Caption) + nRedondeoITF, CCur(Me.LblItf.Caption))
            Set oMov = Nothing
        End If
        'RECUPERACION GRABAR*************************
        Call loGrabarDesem.nSubastaVentaCredPignoraticioSINSubasta(sCredAdjCod, "0000", lsFechaHoraGrab, lsMovNro, nCredAdjnValdeudaTot, _
                                                                   nCredAdjnValdeudaTot + Val(LblItf.Caption), 0, 0, _
                                                                   nCredAdj14k, nCredAdj16k, nCredAdj18k, nCredAdj21k, False, "", sPersCod, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, _
                                                                   CCur(Val(Me.LblItf.Caption)), 0, True, lnMovNro)
        
        If gITF.gbITFAplica Then
            Set oMov = New COMDMov.DCOMMov
            Call oMov.InsertaMovRedondeoITF(lsMovNro, 1, CCur(Val(Me.LblItf.Caption)) + nRedondeoITF, CCur(Val(Me.LblItf.Caption)))
        End If
        Set oMov = Nothing
        Screen.MousePointer = 0
        '********************************************
        lsCadImp = loColImp.nPrintReciboDesembolso(AXDesCon.FechaVencimiento, lsCuenta, lnSaldoCap, lsFechaHoraPrend, _
                           lnMontoEntregar, lnInteresComp, gsNomAge, gsCodUser, CDbl(LblItf.Caption), gImpresora, False, "", "", gbImpTMU)
    
        lcTextImp = "Desea Imprimir las boletas de desembolsos"
        Do
            OptBt2 = MsgBox(lcTextImp, vbInformation + vbYesNo, "Aviso")
            If vbYes = OptBt2 Then
                lcTextImp = "Desea Reimprimir las boletas de desembolsos"
                MsgBox "Cambie de Papel para imprimir las boletas de desembolsos", vbExclamation, "Aviso"
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
                Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
                Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
                Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
                Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita
                Print #nFicSal, lsCadImp & Chr$(12)
                Print #nFicSal, ""
                Close #nFicSal
            End If
        Loop Until OptBt2 = vbNo
        'RECUPERACION
        If MsgBox(" Desea realizar impresión de Recibo ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            Dim oImpColP As COMNColoCPig.NCOMColPImpre
            'Dim lsCadImp As String
            Set oImpColP = New COMNColoCPig.NCOMColPImpre
                lsCadImp = oImpColP.ImpRecupSub(sCredAdjCod, "", lnMontoEntregar + Val(lblMontoNewCred.Caption), lnMontoEntregar + Val(lblMontoNewCred.Caption) + Val(LblItf.Caption), CCur(Val(Me.LblItf.Caption)), gsNomAge, gdFecSis, gsCodUser, gImpresora)
            Set oImpColP = Nothing
            
            Dim loPrevio As previo.clsprevio
            Set loPrevio = New previo.clsprevio
                loPrevio.PrintSpool sLpt, lsCadImp, False
            Set loPrevio = Nothing
            
            Do While True
                If MsgBox("Desea reimprimir ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    Set loPrevio = New previo.clsprevio
                        loPrevio.PrintSpool sLpt, lsCadImp, False
                    Set loPrevio = Nothing
                Else
                    Exit Do
                End If
            Loop
        End If
        
        Set loGrabarDesem = Nothing
        Set loColImp = Nothing
    Else
        MsgBox " Grabación cancelada ", vbInformation, " Aviso "
    End If
    Call LimpiarFormulario
    Exit Sub
ControlError:       ' Rutina de control de errores.
        MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtCtaCMAC.Text = "109"
    txtCtaAGE.Text = gsCodAge
    txtCtaPRD.Text = "705"
End Sub

Private Sub txtCtaNUM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCtaNUM.Text <> "" Then
        sCtaCod = IIf(sCtaCod = "", txtCtaCMAC & txtCtaAGE & txtCtaPRD & txtCtaNUM, sCtaCod)
        Call BuscaContrato(sCtaCod)
    End If
    KeyAscii = SoloNumero(KeyAscii)
End Sub

Public Sub CargarDatosCreditoAdjudicado(ByVal psCtaCod As String)
    Dim loMuestraContrato As New COMDColocPig.DCOMColPContrato
    Dim lrDatosAdjudicacion As New Recordset
    Dim lrDatosPesoOro As New Recordset
    Set lrDatosAdjudicacion = loMuestraContrato.ObtieneValoresAdjudicacion(psCtaCod)
    Set lrDatosPesoOro = loMuestraContrato.ObtieneValorPesoNetoJoya(psCtaCod)
    If Not (lrDatosAdjudicacion.EOF And lrDatosAdjudicacion.BOF) Then
        nCredAdjSaldCap = lrDatosAdjudicacion!nSaldoCap
        nCredAdjnIntComp = lrDatosAdjudicacion!nIntComp
        nCredAdjnIntCompVenc = lrDatosAdjudicacion!nIntCompVenc
        nCredAdjnIntMora = lrDatosAdjudicacion!nIntMora
        nCredAdjnCostoNotif = lrDatosAdjudicacion!nCostoNotif
        nCredAdjnValTasacion = lrDatosAdjudicacion!nValTasacion
        nCredAdjnValdeudaTot = nCredAdjSaldCap + nCredAdjnIntComp + nCredAdjnIntCompVenc + nCredAdjnIntMora + nCredAdjnCostoNotif
    End If
    If Not (lrDatosPesoOro.EOF And lrDatosPesoOro.BOF) Then
        Dim i As Integer
        For i = 0 To lrDatosPesoOro.RecordCount - 1
            Select Case lrDatosPesoOro!cKilataje
                Case 14
                    nCredAdj14k = lrDatosPesoOro!nPesoOro
                Case 16
                    nCredAdj16k = lrDatosPesoOro!nPesoOro
                Case 18
                    nCredAdj18k = lrDatosPesoOro!nPesoOro
                Case 21
                    nCredAdj21k = lrDatosPesoOro!nPesoOro
            End Select
            lrDatosPesoOro.MoveNext
        Next
    End If
End Sub
Public Function SoloNumero(ByVal pKeyAscii As Integer)
    Select Case pKeyAscii
        Case 48 To 57
            SoloNumero = pKeyAscii
        Case Else
            SoloNumero = 0
    End Select
End Function

Public Sub LimpiarFormulario()
    txtCtaCMAC.Text = "109"
    txtCtaAGE.Text = gsCodAge
    txtCtaPRD.Text = "705"
    txtCtaNUM.Text = ""
    cmdBuscarCliente.Enabled = True
    AXDesCon.Limpiar
    lblMontoNewCred.Caption = ""
    lblInicial.Caption = ""
    LblItf.Caption = ""
    lblTotalCobrar.Caption = ""
    lblCtaCMAC.Caption = ""
    lblCtaAGE.Caption = ""
    lblCtaPRD.Caption = ""
    lblCtaNum.Caption = ""
    cmdgrabar.Enabled = True
End Sub

