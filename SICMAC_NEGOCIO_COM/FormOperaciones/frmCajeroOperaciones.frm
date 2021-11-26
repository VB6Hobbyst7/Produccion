VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCajeroOperaciones 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6585
   ClientLeft      =   3990
   ClientTop       =   2430
   ClientWidth     =   8415
   Icon            =   "frmCajeroOperaciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   375
      TabIndex        =   3
      Top             =   1695
      Width           =   1020
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   375
      TabIndex        =   2
      Top             =   1230
      Width           =   1020
   End
   Begin VB.Frame fraOperacion 
      Caption         =   "Digite Operación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   960
      Left            =   90
      TabIndex        =   5
      Top             =   60
      Width           =   1755
      Begin VB.TextBox txtOperacion 
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
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   240
         TabIndex        =   0
         Top             =   390
         Width           =   1260
      End
   End
   Begin VB.Frame fraOperaciones 
      Caption         =   "Seleccione Operación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   6435
      Left            =   1920
      TabIndex        =   4
      Top             =   60
      Width           =   6435
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   1260
         Top             =   5700
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCajeroOperaciones.frx":030A
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCajeroOperaciones.frx":065C
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCajeroOperaciones.frx":09AE
               Key             =   "Hijito"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCajeroOperaciones.frx":0D00
               Key             =   "Bebe"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwOperacion 
         Height          =   6075
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   10716
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCajeroOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sOpePadre As String, sOpeHijo As String, sOpeHijito As String

Dim Index200000 As Integer ' *** MAVM:Auditoria
Dim Index201000 As Integer ' *** MAVM:Auditoria
Dim Index201002 As Integer ' *** MAVM:Auditoria
Dim gsOpeCodTemp As String '*** PEAC 20140318

Public Sub Inicia(ByVal sCaption As String, ByVal rsUsu As ADODB.Recordset)

Me.Caption = sCaption
tvwOperacion.Nodes.Clear
LlenaArbol rsUsu
Me.Show '1
End Sub

Private Sub LlenaArbol(ByVal rsUsu As ADODB.Recordset)
Dim sOperacion As String, sOpeCod As String
Dim nodOpe As Node
Dim i As Integer ' ***MAVM:Auditoria

On Error GoTo ErrCab

Do While Not rsUsu.EOF
    sOpeCod = rsUsu("cOpeCod")
    sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
    Select Case rsUsu("nOpeNiv")
        Case "1"
            sOpePadre = "P" & sOpeCod
            Set nodOpe = tvwOperacion.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
            nodOpe.Tag = sOpeCod
        Case "2"
            sOpeHijo = "H" & sOpeCod
            Set nodOpe = tvwOperacion.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
            nodOpe.Tag = sOpeCod
        Case "3"
            sOpeHijito = "J" & sOpeCod
            Set nodOpe = tvwOperacion.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
            nodOpe.Tag = sOpeCod
        Case "4"
            Set nodOpe = tvwOperacion.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
            nodOpe.Tag = sOpeCod
    End Select
    i = i + 1 ' ***MAVM:Auditoria
    If sOpeCod = "200000" Then Index200000 = i ' ***MAVM:Auditoria
    If sOpeCod = "201000" Then Index201000 = i ' ***MAVM:Auditoria
    If sOpeCod = "201002" Then Index201002 = i ' ***MAVM:Auditoria
    rsUsu.MoveNext
Loop
Exit Sub
ErrCab:
    If err.Number = 35603 Or err.Number = 424 Or err.Number = 35601 Then
        MsgBox "Falta Asignar Permiso a la Cabecera de la Operacion : " & sOperacion, vbCritical, "Aviso"
    Else
        MsgBox err.Description, vbCritical, "Aviso"
    End If
    If rsUsu.State = 1 Then
        rsUsu.Close
        Set rsUsu = Nothing
    End If

End Sub

'Private Sub EjecutaOperacion(ByVal nOperacion As CaptacOperacion, ByVal sDescOperacion As String)
'    Dim asd As CaptacOperacion, cNroProceso As String, clscol As COMNColoCPig.NCOMColPRecGar
'    Dim oCaj As COMNCajaGeneral.NCOMCajero
'    gsOpeCod = Trim(nOperacion)
'    gsOpeDesc = Trim(sDescOperacion)
'
'    Set oCaj = New COMNCajaGeneral.NCOMCajero
'    If oCaj.YaRealizoCierreAgencia(gsCodAge, gdFecSis) Then
'        'RECO20151111 ERS061-2015******************
'        If Not VerificaGrupoPermisoPostCierre Then
'            MsgBox "Ya se realizó el cierre de caja de la agencia. No es posible realizar transacciones", vbExclamation, "Aviso"
'            Set oCaj = Nothing
'            Unload Me
'            Exit Sub
'        End If
'        'RECO FIN *********************************
'    End If
'
'    Select Case nOperacion
'    'Aperturas
'
'        Case gAhoApeEfec, gAhoApeChq, gAhoApeTransf, gAhoApeCargoCta 'JUEZ 20131212 Se agregó gAhoApeCargoCta
'            frmCapAperturas.Inicia gCapAhorros, nOperacion, , sDescOperacion
'        Case gAhoApeLoteEfec, gAhoApeLoteChq, gAhoApeLoteTransfBanco 'RIRO20140407 ERS017 Se agrego gAhoApeLoteTransfBanco
'            frmCapAperturasLote.Inicia gCapAhorros, nOperacion, sDescOperacion
'        Case gPFApeEfec, gPFApeChq, gPFApeTransf, gPFApeCargoCta 'JUEZ 20131212 Se agregó gPFApeCargoCta
'            frmCapAperturas.Inicia gCapPlazoFijo, nOperacion, , sDescOperacion
'        Case gPFApeLoteEfec, gPFApeLoteChq
'            frmCapAperturasLote.Inicia gCapPlazoFijo, nOperacion, sDescOperacion
'        Case gCTSApeEfec, gCTSApeChq, gCTSApeTransf
'            frmCapAperturas.Inicia gCapCTS, nOperacion, , sDescOperacion
'        Case gCTSApeLoteEfec, gCTSApeLoteChq
'            frmCapAperturasLote.Inicia gCapCTS, nOperacion, sDescOperacion
'        Case gServGiroApertEfec
'            frmGiroApertura.Show 1
'
'
'    'Abonos
'
'    '****OTROS ABONOS A CUENTAS
'        Case 200242, 200247, 200248, 200249, 200250, 200251
'            frmCapAbonos.Inicia gCapAhorros, nOperacion, , , sDescOperacion
'        Case 200252
'            frmCapAbonos.Inicia gCapAhorros, nOperacion, , , sDescOperacion
'        Case 200253
'            frmCapAbonos.Inicia gCapAhorros, nOperacion, , , sDescOperacion
'        ' RIRO20131102 SEGUN TI-ERS145-2013
'        Case gAhoDepDirectoClub
'            frmDepositoCuentaClub.Show 1
'
'
'    '****
'
'        Case gAhoDepEfec, gAhoDepChq, gAhoDepTransf, gAhoDepDevFondoGar, _
'            gAhoDepAboOtrosConceptos, gAhoDepOtrosIngRRHH, gAhoDepDevCredPersonales, "200243", "200244", "200245"
'            frmCapAbonos.Inicia gCapAhorros, nOperacion, , , sDescOperacion
'        'By capi 05032009 Acta 025-2009
'        'Case gCTSDepEfec, gCTSDepChq, gCTSDepTransf
'        Case gCTSDepEfec, gCTSDepChq, gCTSDepTransf, gCTSDepAboOtrosConceptos
'            frmCapAbonos.Inicia gCapCTS, nOperacion, , , sDescOperacion
'        '***Agregado por ELRO el 20121116, según OYP-RFC101-2012
'        Case gCTSDepLotEfec, gCTSDepLotChq, gCTSDepLotTransf
'            frmCapDepositosLote.iniciarFormulario gCapCTS, nOperacion, "Deposito en Lote"
'        '***Fin Agregado por ELRO el 20121116*******************
'        Case gAhoDepEntConv
'            'frmCapServConvenioOpe.Show 1
'        Case gPFAumCapEfec, gPFAumCapchq, gPFAumCapTasaPactEfec, gPFAumCapTasaPactChq, gPFDismCapEfec, gPFAumCapTasaPactTrans, gPFAumCapTrans, gPFAumCapCargoCta 'JUEZ 20131212 Se agregó gPFAumCapCargoCta
'            frmCapOpePlazoFijo.Inicia nOperacion, sDescOperacion
'        Case gAhoDepCtaRecaudoEcotaxi 'EJVG20120515
'            frmRecaudoEcotaxi.Inicio
'    'Cargos
'
'    '*****OTROS CARGOS A CUENTAS
'        Case 200331 To 200357
'            frmCapCargos.Inicia gCapAhorros, nOperacion, sDescOperacion
'    '*******
'
'        Case gAhoRetEfec, gAhoRetOP, gAhoRetTransf, gAhoRetOPCanje, gAhoRetEmiChq, gAhoRetEmiChqCanjeOP, _
'            gAhoRetRetencionJudicial, gAhoRetDuplicadoTarj, gAhoRetComOrdPagDev, gAhoRetConsultaSaldos, _
'            gAhoRetPorteCargoCuenta, gAhoRetComVentaBases, gAhoRetComTransferencia, gAhoRetChequeDevuelto, _
'            gAhoRetOtrosConceptos
'            frmCapCargos.Inicia gCapAhorros, nOperacion, sDescOperacion
'        Case gAhoRetFondoFijo, gAhoRetFondoFijoCanje
'            frmCapFondoFijo.Inicia nOperacion, sDescOperacion
'        Case gCTSRetEfec, gCTSRetTransf, "220303"
'            frmCapCargos.Inicia gCapCTS, nOperacion, sDescOperacion
'
'    'Cancelaciones
'        Case gAhoCancAct, gAhoCancTransfAct, gAhoCancTransfAbCtaBco ' RIRO20131212 ERS137
'            frmCapCancelacion.Inicia gCapAhorros, nOperacion, sDescOperacion
'        Case gPFCancEfec, gPFCancTransf
'            frmCapOpePlazoFijo.Inicia nOperacion, sDescOperacion
'        Case gCTSCancEfec, gCTSCancTransf, gCTSCancTransfBco ' RIRO20131226 ERS137
'            frmCapCancelacion.Inicia gCapCTS, nOperacion, sDescOperacion
'        Case gServGiroCancEfec
'            frmGiroCancelacion.Show 1
'        'RECO20140415 ERS008-2014 *****************************************
'        Case 310401
'            frmGiroMantDestinatario.Show 1
'        Case 310402
'            frmGiroAnulacion.Show 1
'        Case 310501
'            frmGiroMovimiento.Show 1
'        'RECO FIN *********************************************************
'        'Transferencias
'        Case gAhoTransferencia
'            'frmCapTransferencia.Show 1
'            'frmCapTransferenciaCambios.Show 1
'            frmCapTransferenciaCambios.Inicia
'        Case gAhoTransAbonoL 'Transferencia en LOTE GITU 15-10-2012
'            frmCapTransferenciaCambiosLote.Inicia
'        'Case gAhoOperacionesPendientes
'    '       frmCapAutorizacion.Inicio
'    'Consulta de Saldos
'        Case gAhoConsSaldo
'             frmCapConsultaSaldos.Inicia gCapAhorros
'        Case gPFConsSaldo
'            frmCapConsultaSaldos.Inicia gCapPlazoFijo
'        Case gCTSConsSaldo
'            frmCapConsultaSaldos.Inicia gCapCTS
'        'Consulta de Movimientos
'        Case gAhoConsMovimiento
'            frmCapConsultaMovimientos.Inicia gCapAhorros, gAhoConsMovimiento
'        Case gCTSConsMovimiento
'            frmCapConsultaMovimientos.Inicia gCapCTS, gCTSConsMovimiento
'        Case gPFConsMovimiento
'            frmCapConsultaMovimientos.Inicia gCapPlazoFijo, gPFConsMovimiento
'        'Retiro de Intereses
'        Case gPFRetInt, gPFRetIntAboAho, gPFRetIntAdelantado, gPFRetIntAboCtaBanco 'RIRO20131212 ERS137
'            frmCapOpePlazoFijo.Inicia nOperacion, sDescOperacion
'
'    'Duplicado de Certificado de Plazo Fijo
'        Case gPFDupCert
'            frmCapDupCertPF.Show 1
'        Case gPFBusqCredOend
'            'frmCapBusqCredPendPF.Show 1
'    'Migracion
'    '***Agregado por ELRO el 20130327, según TI-ERS011-2013****
'        Case gAhoMigracion
'            frmCapMigracionAhorros.Show 1
'    '***Fin Agregado por ELRO el 20130327, según TI-ERS011-2013
'        'Compra Venta
'        Case gOpeCajeroMECompra
'            'frmCajeroCompraVenta.Show 1
'            frmCompraVenta.Show 1
'        Case gOpeCajeroMEVenta
'            'frmCajeroCompraVenta.Show 1
'            frmCompraVenta.Show 1
'
'        Case gOpeCajeroMECompraEsp
'            frmCajeroCompraVentaEsp.Show 1
'        Case gOpeCajeroMEVentaEsp
'            frmCajeroCompraVentaEsp.Show 1
'
'        'Extorno Compra - Venta
'        Case gOpeCajeroMEExtCompra, gOpeCajeroMEExtVenta, gOpeCajeroMEExtCompraEsp, gOpeCajeroMEExtVentaEsp
'            frmCajeroExtornos.Show 1
'
'        'Boveda Agencia
'        Case gOpeBoveAgeConfHabCG
'            'frmCajaGenLista.Show 1
'            frmRemesaConfirmacion.Show 1 'EJVG20140905
'        'Case gOpeBoveAgeHabAgeACG
'            'frmCajaGenHabilitacion.Show 1
'        Case gOpeBoveAgeHabEntreAge
'            'frmCajaGenHabilitacion.Show 1
'            frmRemesaAgenciaToIFiAgencia.Show 1 'EJVG20140905
'        Case gOpeBoveAgeHabCajero
'            frmCajeroHab.Show 1
'        Case gOpeBoveAgeExtConfHabCG
'            frmRemesaConfirmacionExt.Show 1 'EJVG20140905
'        'Case gOpeBoveAgeExtConfHabCG, gOpeBoveAgeExtHabAgeACG, gOpeBoveAgeExtHabEntreAge
'            'frmCajaGenLista.Show 1
'        Case gOpeBoveAgeExtHabEntreAge
'            frmRemesaAgenciaToIFiAgenciaExt.Show 1 'EJVG20140905
'        Case gOpeBoveAgeExtHabCajero
'            frmCajeroExtornos.Inicia "BOVEDA DE AGENCIA " & sDescOperacion
'        Case gOpeBoveAgeConfDevCaj
'            frmCajeroExtornos.Inicia "BOVEDA DE AGENCIA " & sDescOperacion
'        Case gOpeBoveAgeRegEfect
'            frmCajaGenEfectivo.RegistroEfectivo True, gOpeBoveAgeRegEfect         'MADM 20110211
'        Case gOpeBoveAgeExtRegEfect ' DAOR 20080204
'            frmCajeroExtornos.Inicia "BOVEDA DE AGENCIA " & sDescOperacion
'        Case gOpeBoveAgeRegSobFalt
'            If gbVerificaRegistroEfectivo Then
'                Set oCaj = New COMNCajaGeneral.NCOMCajero
'                If oCaj.YaRegistroEfectivo(gsCodAge, gdFecSis, gsUsuarioBOVEDA, gOpeBoveAgeRegEfect) Then
'                    Set oCaj = Nothing
'                    frmCajeroIngEgre.Inicia True, False
'                Else
'                    Set oCaj = Nothing
'                    MsgBox "Aun No Se ha realizado el Eegistro de Efectivo de Bóveda.", vbInformation, "Aviso"
'                End If
'            Else
'                frmCajeroIngEgre.Inicia True, False
'            End If
'        Case gOpeBoveAgeExtRegSobFalt
'            frmCajeroExtornos.Inicia "BOVEDA DE AGENCIA " & sDescOperacion
'
''    Case 121000
''        frmPigAmortizacion.Inicio nOperacion, sDescOperacion, "", ""
'
'        'Operaciones Cajero
'        Case gOpeHabCajRegEfect
'        'MIOL 20120601, SEGUN RQ12093 *****************************************************************
'        'MADM 20110203
'        '    frmCajaGenEfectivo.RegistroEfectivo
'        '    frmCajeroBilletajePre.Show 1
'             frmCajaGenEfectivo.RegistroEfectivo True, gOpeHabCajRegEfect
'        'END MADM
'        'END MIOL *************************************************************************************
'        'MADM 20110926
'        Case 901035
'            frmCajeroExtornos.Inicia "CAJERO " & sDescOperacion
'        'END MADM
'        Case gOpeHabCajDevABove
'            frmCajeroHab.Show 1
'        Case gOpeHabCajTransfEfectCajeros
'            frmCajeroHab.Show 1
'        Case gOpeHabCajConfHabBovAge
'            frmCajeroExtornos.Inicia "CAJERO " & sDescOperacion
'        Case gOpeHabCajRegSobFalt
'            If gbVerificaRegistroEfectivo Then
'                Set oCaj = New COMNCajaGeneral.NCOMCajero
'                If oCaj.YaRegistroEfectivo(gsCodAge, gdFecSis, gsCodUser, gOpeHabCajRegEfect) Then
'                    Set oCaj = Nothing
'                    frmCajeroIngEgre.Inicia False, False
'                Else
'                    Set oCaj = Nothing
'                    MsgBox "Aun No Se ha realizado el Eegistro de Efectivo de Cajero.", vbInformation, "Aviso"
'                End If
'            Else
'                frmCajeroIngEgre.Inicia False, False
'            End If
'        Case gOpeHabCajIngEfectRegulaFalt
'            'MAVM 20120328 ***
'            'frmRegularizaSobFal.Ini sDescOperacion
'            frmRegularizaSobFal.Ini sDescOperacion, nOperacion
'            '***
'        Case gOpeHabCajExtTransfEfectCajeros
'            frmCajeroExtornos.Inicia "CAJERO " & sDescOperacion
'        Case gOpeHabCajExtConfHabBovAge, gOpeHabCajExtIngEfectRegulaFalt, gOpeHabCajExtRegSobFalt
'            frmCajeroExtornos.Inicia "CAJERO " & sDescOperacion
'        Case gOpeHabCajExtDevABove, gOpeHabCajExtDevBilletaje
'            frmCajeroExtornos.Inicia "CAJERO " & sDescOperacion
'
'        'Operaciones con Cheques
'        Case gChqOpeRegistro
'            'frmIngCheques.Inicio True, Trim(nOperacion), True, 0, gMonedaNacional, , , 0, False, "", True
'        'EJVG20140408 ***
'            frmCheque.Registrar
'        Case gChqOpeMantenimiento
'            frmChequeEditSel.Show 1
'        Case gChqOpeDeposito
'            frmChequeDeposito.Show 1
'        Case 900035
'            frmChequeOpePendiente.Show 1
'        Case gChqOpeExtRegistro
'            frmChequeExtorno.Show 1
'        Case gChqOpeExtDeposito
'            frmChequeDepositoExtorno.Show 1
'        'END EJVG *******
'        Case gChqOpeModFechaValor, gChqOpeValorInmediata, gChqOpeConsultaEstado, _
'            gChqOpeExtValorInmediata
'            'frmChqMantenimiento.Inicia nOperacion, sDescOperacion
'        'MIOL 20130511, SEGUN RQ13251 ***************
'        Case gChqOpeMantGirador
'            'frmCambioGirador.Show 1
'        'END MIOL ***********************************
'        '***Agregado por ELRO el 20120627, según OYP-RFC024-2012
'        Case gVouOpeRegistro
'            'frmCapRegVouDep.Show 1
'            frmCapRegVouDep_NEW.Nuevo 'EJVG20130903
'        Case gVouOpeEditar
'            frmCapRegVouDepEdi.Show 1
'        '***Fin Agregado por ELRO*******************************
'        '***Agregado por ELRO el 20130712, según RFC1306270002
'        Case gCapConSerPagDeb
'            frmCapServicioPagoDebito.Show 1
'        '***Fin Agregado por ELRO el 20130712, según RFC1306270002
'        'FRHU 20141201 ERS048-2014
'        Case gCapNotaDeCargo
'             frmOpeNotaAbonoCargo.Inicio nOperacion
'        Case gCapNotaDeAbono
'             frmOpeNotaAbonoCargo.Inicio nOperacion
'        'FIN FRHU 20141201
'        '*** PEAC 20081002
'        '*** SE AGREGO UN PARAMERO (nOperacion) A TODAS LAS LLAMADAS A: frmCapExtornos.Inicia gAhoDepTransf, sDescOperacion, gCapAhorros, nOperacion
'        '*** PARA PODER VALIDAR EL VISTO ELECTRONICO
'
'        'Extornos de Captaciones
'        'Extornos de Aperturas
'        Case gAhoExtApeEfec
'            frmCapExtornos.Inicia gAhoApeEfec, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtApeChq
'            frmCapExtornos.Inicia gAhoApeChq, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtApeTransf
'            frmCapExtornos.Inicia gAhoApeTransf, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtApeLoteEfec
'            frmCapExtornos.Inicia gAhoApeLoteEfec, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtApeLoteChq
'            frmCapExtornos.Inicia gAhoApeLoteChq, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtApeCargoCta 'JUEZ 20131226
'            frmCapExtornos.Inicia gAhoApeCargoCta, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtApeLoteTransfBanco 'RIRO20140530 ERS017
'            frmCapExtornos.Inicia gAhoApeLoteTransfBanco, sDescOperacion, gCapAhorros, nOperacion
'        Case gPFExtApeChq
'            frmCapExtornos.Inicia gPFApeChq, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtApeTransf
'            frmCapExtornos.Inicia gPFApeTransf, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtApeEfec
'            frmCapExtornos.Inicia gPFApeEfec, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtApeLoteEfec
'            frmCapExtornos.Inicia gPFApeLoteEfec, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtApeLoteChq
'            frmCapExtornos.Inicia gPFApeLoteChq, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtApeCargoCta 'JUEZ 20131226
'            frmCapExtornos.Inicia gPFApeCargoCta, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gCTSExtApeChq
'            frmCapExtornos.Inicia gCTSApeChq, sDescOperacion, gCapCTS, nOperacion
'        Case gCTSExtApeTransf
'            frmCapExtornos.Inicia gCTSApeTransf, sDescOperacion, gCapCTS, nOperacion
'        Case gCTSExtApeEfec
'            frmCapExtornos.Inicia gCTSApeEfec, sDescOperacion, gCapCTS, nOperacion
'        Case gCTSExtApeLoteEfec
'            frmCapExtornos.Inicia gCTSApeLoteEfec, sDescOperacion, gCapCTS, nOperacion
'        Case gCTSExtApeLoteChq
'            frmCapExtornos.Inicia gCTSApeLoteChq, sDescOperacion, gCapCTS, nOperacion
'
'        Case gAhoExtDepEfec
'            frmCapExtornos.Inicia gAhoDepEfec, sDescOperacion, gCapAhorros, nOperacion
'        Case "230207"
'            frmCapExtornos.Inicia 200207, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtDepChq
'            frmCapExtornos.Inicia gAhoDepChq, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtDepTransf
'            frmCapExtornos.Inicia gAhoDepTransf, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtDepPagServEdelnor
'            frmCapExtornos.Inicia gAhoDepPagServEdelnor, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtDepOtrosConceptos
'            frmCapExtornos.Inicia gAhoDepAboOtrosConceptos, sDescOperacion, gCapAhorros, nOperacion
'        Case "230209"
'            frmCapExtornos.Inicia 200209, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtDepDevCredPersonales
'            frmCapExtornos.Inicia gAhoDepDevCredPersonales, sDescOperacion, gCapAhorros, nOperacion
'        Case "230243"
'            frmCapExtornos.Inicia "200243", sDescOperacion, gCapAhorros, nOperacion
'        Case "230244"
'            frmCapExtornos.Inicia "200244", sDescOperacion, gCapAhorros, nOperacion
'        Case "230245"
'            frmCapExtornos.Inicia "200245", sDescOperacion, gCapAhorros, nOperacion
'        Case "230246"
'            frmCapExtornos.Inicia "200246", sDescOperacion, gCapAhorros, nOperacion
'        Case "230247"
'            frmCapExtornos.Inicia "200247", sDescOperacion, gCapAhorros, nOperacion
'        Case "230248"
'            frmCapExtornos.Inicia "200248", sDescOperacion, gCapAhorros, nOperacion
'        Case "230249"
'            frmCapExtornos.Inicia "200249", sDescOperacion, gCapAhorros, nOperacion
'        Case "230250"
'            frmCapExtornos.Inicia "200250", sDescOperacion, gCapAhorros, nOperacion
'        Case "230251"
'            frmCapExtornos.Inicia "200251", sDescOperacion, gCapAhorros, nOperacion
'        Case "230252"
'            frmCapExtornos.Inicia "200204", sDescOperacion, gCapAhorros, nOperacion
'        Case "230254"
'            frmCapExtornos.Inicia "200252", sDescOperacion, gCapAhorros, nOperacion
'        Case "230255"
'            frmCapExtornos.Inicia "200253", sDescOperacion, gCapAhorros, nOperacion
'        Case gCMACOAAhoExtDepEfec
'            frmCapExtornos.Inicia gCMACOAAhoDepEfec, sDescOperacion, gCapAhorros, nOperacion
'
'        ' RIRO20131102 ERS145
'        Case gAhoExtDirectoClub
'            frmCapExtornos.Inicia gAhoDepDirectoClub, sDescOperacion, gCapAhorros, nOperacion
'        ' FIN RIRO
'        ' *** RIRO 20140530 ERS017
'        Case gAhoExtDepositoHaberesEnLoteEfec
'            frmCapExtornos.Inicia gAhoDepositoHaberesEnLoteEfec, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtDepositoHaberesEnLoteTransf
'            frmCapExtornos.Inicia gAhoDepositoHaberesEnLoteTransf, sDescOperacion, gCapAhorros, nOperacion
'        ' *** END RIRO
'        Case gCMACOAAhoExtDepChq
'            frmCapExtornos.Inicia gCMACOAAhoDepChq, sDescOperacion, gCapAhorros, nOperacion
'        Case gCTSExtDepEfec
'            frmCapExtornos.Inicia gCTSDepEfec, sDescOperacion, gCapCTS, nOperacion
'        Case gCTSExtDepChq
'            frmCapExtornos.Inicia gCTSDepChq, sDescOperacion, gCapCTS, nOperacion
'        Case gCTSExtDepTransf
'            frmCapExtornos.Inicia gCTSDepTransf, sDescOperacion, gCapCTS, nOperacion
'        Case gCMACOACTSExtDepEfec
'            frmCapExtornos.Inicia gCMACOACTSDepEfec, sDescOperacion, gCapCTS, nOperacion
'        '***Agregado por ELRO el 20121120, según OYP-RFC101-2012
'        Case gCTSExtDepLotEfec
'            frmCapExtornos.Inicia gCTSDepLotEfec, sDescOperacion, gCapCTS, nOperacion
'        'Case gCTSExtDepLotChq
'            'frmCapExtornos.Inicia gCTSDepLotChq, sDescOperacion, gCapCTS, nOperacion
'        'Case gCTSExtDepLotTransf
'            'frmCapExtornos.Inicia gCTSDepLotTransf, sDescOperacion, gCapCTS, nOperacion
'        '***Fin Agregado por ELRO el 20121120*******************
'        Case gAhoExtRetEfec
'            frmCapExtornos.Inicia gAhoRetEfec, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetRetencionJudicial
'            frmCapExtornos.Inicia gAhoRetRetencionJudicial, sDescOperacion, gCapAhorros, nOperacion
'
'        Case gAhoExtRetDuplicadoTarj
'            frmCapExtornos.Inicia gAhoRetDuplicadoTarj, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetOP
'            frmCapExtornos.Inicia gAhoRetOP, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetTransf
'            frmCapExtornos.Inicia gAhoRetTransf, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetOPCanje
'            frmCapExtornos.Inicia gAhoRetOPCanje, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetOPCert
'            frmCapExtornos.Inicia gAhoRetOPCert, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetFondoFijo
'            frmCapExtornos.Inicia gAhoRetFondoFijo, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetOPCertCanje
'            frmCapExtornos.Inicia gAhoRetOPCertCanje, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetAnulChq
'            frmCapExtornos.Inicia gAhoRetAnulChq, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetComServEdelnor
'            frmCapExtornos.Inicia gAhoRetComServEDELNOR, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetChequeDevuelto
'            frmCapExtornos.Inicia gAhoRetChequeDevuelto, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetComTransferencia
'            frmCapExtornos.Inicia gAhoRetComTransferencia, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetEmiChq
'            frmCapExtornos.Inicia gAhoRetEmiChq, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetComOrdPagDev
'            frmCapExtornos.Inicia gAhoRetComOrdPagDev, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetEmiChqCanjeOP
'            frmCapExtornos.Inicia gAhoRetEmiChqCanjeOP, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetOtrosConceptos
'            frmCapExtornos.Inicia gAhoRetOtrosConceptos, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetConsultaSaldos
'            frmCapExtornos.Inicia gAhoRetConsultaSaldos, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetPorteCargoCuenta
'            frmCapExtornos.Inicia gAhoRetPorteCargoCuenta, sDescOperacion, gCapAhorros, nOperacion
'
'        Case 230332
'            frmCapExtornos.Inicia 200331, sDescOperacion, gCapAhorros, nOperacion
'        Case 230333
'            frmCapExtornos.Inicia 200332, sDescOperacion, gCapAhorros, nOperacion
'        Case 230334
'            frmCapExtornos.Inicia 200333, sDescOperacion, gCapAhorros, nOperacion
'        Case 230335
'            frmCapExtornos.Inicia 200334, sDescOperacion, gCapAhorros, nOperacion
'        Case 230336
'            frmCapExtornos.Inicia 200335, sDescOperacion, gCapAhorros, nOperacion
'        Case 230337
'            frmCapExtornos.Inicia 200336, sDescOperacion, gCapAhorros, nOperacion
'        Case 230338
'            frmCapExtornos.Inicia 200337, sDescOperacion, gCapAhorros, nOperacion
'        Case 230339
'            frmCapExtornos.Inicia 200338, sDescOperacion, gCapAhorros, nOperacion
'        Case 230340
'            frmCapExtornos.Inicia 200339, sDescOperacion, gCapAhorros, nOperacion
'        Case 230341
'            frmCapExtornos.Inicia 200340, sDescOperacion, gCapAhorros, nOperacion
'        Case 230342
'            frmCapExtornos.Inicia 200341, sDescOperacion, gCapAhorros, nOperacion
'        Case 230343
'            frmCapExtornos.Inicia 200342, sDescOperacion, gCapAhorros, nOperacion
'        Case 230344
'            frmCapExtornos.Inicia 200343, sDescOperacion, gCapAhorros, nOperacion
'        Case 230345
'            frmCapExtornos.Inicia 200344, sDescOperacion, gCapAhorros, nOperacion
'        Case 230346
'            frmCapExtornos.Inicia 200345, sDescOperacion, gCapAhorros, nOperacion
'        Case 230347
'            frmCapExtornos.Inicia 200346, sDescOperacion, gCapAhorros, nOperacion
'        Case 230348
'            frmCapExtornos.Inicia 200347, sDescOperacion, gCapAhorros, nOperacion
'        Case 230349
'            frmCapExtornos.Inicia 200348, sDescOperacion, gCapAhorros, nOperacion
'        Case 230350
'            frmCapExtornos.Inicia 200349, sDescOperacion, gCapAhorros, nOperacion
'        Case 230351
'            frmCapExtornos.Inicia 200350, sDescOperacion, gCapAhorros, nOperacion
'        Case 230352
'            frmCapExtornos.Inicia 200351, sDescOperacion, gCapAhorros, nOperacion
'        Case 230353
'            frmCapExtornos.Inicia 200352, sDescOperacion, gCapAhorros, nOperacion
'        Case 230354
'            frmCapExtornos.Inicia 200353, sDescOperacion, gCapAhorros, nOperacion
'        Case 230355
'            frmCapExtornos.Inicia 200354, sDescOperacion, gCapAhorros, nOperacion
'        Case 230356
'            frmCapExtornos.Inicia 200355, sDescOperacion, gCapAhorros, nOperacion
'        Case 230357
'            frmCapExtornos.Inicia 200356, sDescOperacion, gCapAhorros, nOperacion
'        Case 230358
'            frmCapExtornos.Inicia 200357, sDescOperacion, gCapAhorros, nOperacion
'        Case 230224
'            frmCapExtornos.Inicia 200224, sDescOperacion, gCapAhorros, nOperacion
'        Case 230360
'            frmCapExtornos.Inicia 200601, sDescOperacion, gCapAhorros, nOperacion
'
'        Case gCMACOAAhoExtRetEfec
'            frmCapExtornos.Inicia gCMACOAAhoRetEfec, sDescOperacion, gCapAhorros, nOperacion
'        Case gCMACOAAhoExtRetOP
'            frmCapExtornos.Inicia gCMACOAAhoRetOP, sDescOperacion, gCapAhorros, nOperacion
'        Case gCMACOAAhoExtRetOPCert
'            frmCapExtornos.Inicia gCMACOAAhoRetOPCert, sDescOperacion, gCapAhorros, nOperacion
'        Case gCTSExtRetEfec
'            frmCapExtornos.Inicia gCTSRetEfec, sDescOperacion, gCapCTS, nOperacion
'        Case gCTSExtRetTransf
'            frmCapExtornos.Inicia gCTSRetTransf, sDescOperacion, gCapCTS, nOperacion
'        Case "250303"
'            frmCapExtornos.Inicia "220303", sDescOperacion, gCapCTS, nOperacion
'        Case gCTSExtCargoComDivAho 'JUEZ 20130906
'            frmCapExtornos.Inicia gCTSCargoCobroComDiversasAho, sDescOperacion, gCapCTS, nOperacion
'        Case gCMACOACTSExtRetEfec
'            frmCapExtornos.Inicia gCMACOACTSRetEfec, sDescOperacion, gCapCTS, nOperacion
'        Case gAhoExtRetEmiChq
'            frmCapExtornos.Inicia gAhoRetEmiChq, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtRetEmiChqCanjeOP
'            frmCapExtornos.Inicia gAhoRetEmiChqCanjeOP, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtTransCargo
'            frmCapExtornos.Inicia gAhoTransCargo, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtTransAbono
'            frmCapExtornos.Inicia gAhoTransAbono, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtTransAbonoL
'            frmCapExtornos.Inicia gAhoTransAbonoL, sDescOperacion, gCapAhorros, nOperacion
'        Case gPFExtRetIntCash  '*** PEAC 20091229 esta opcion estaba comentada se habilito GITU 16-12-2010
'            frmCapExtornos.Inicia gPFRetIntAdelantado, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtRetIntAboAho
'            frmCapExtornos.Inicia gPFRetIntAboAho, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtRetIntAboCtaTransf 'RIRO20131212 ERS137
'            frmCapExtornos.Inicia gPFRetIntAboCtaBanco, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtCancTransfAbBco 'RIRO20131212 ERS137
'            frmCapExtornos.Inicia gPFCancTransf, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gCMACOAPFExtRetInt
'            frmCapExtornos.Inicia gCMACOAPFRetInt, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtRetInt
'            frmCapExtornos.Inicia gPFRetInt, sDescOperacion, gCapPlazoFijo, nOperacion
'        'Extornos de Aumento/Disminución de Capital
'        Case gPFExtAumCapEfec
'            frmCapExtornos.Inicia gPFAumCapEfec, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtAumCapTasaPactEfec
'            frmCapExtornos.Inicia gPFAumCapTasaPactEfec, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtAumCapchq
'            frmCapExtornos.Inicia gPFAumCapchq, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtAumCapTasaPactChq
'            frmCapExtornos.Inicia gPFAumCapTasaPactChq, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtAumCapTrans
'             frmCapExtornos.Inicia gPFAumCapTrans, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtAumCapTasaPactTrans
'            frmCapExtornos.Inicia gPFAumCapTasaPactTrans, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtDismCapEfec
'            frmCapExtornos.Inicia gPFDismCapEfec, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtAumCapCargoCta 'JUEZ 20131226
'            frmCapExtornos.Inicia gPFAumCapCargoCta, sDescOperacion, gCapPlazoFijo, nOperacion
'
'        Case gAhoExtCancAct
'            frmCapExtornos.Inicia gAhoCancAct, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtCancTransfAct
'            frmCapExtornos.Inicia gAhoCancTransfAct, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtCanctransf 'RIRO20131212 ERS137
'            frmCapExtornos.Inicia gAhoCancTransfAbCtaBco, sDescOperacion, gCapAhorros, nOperacion
'        Case gPFExtCancEfec
'            frmCapExtornos.Inicia gPFCancEfec, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gPFExtCancTransf
'            frmCapExtornos.Inicia gPFCancTransf, sDescOperacion, gCapPlazoFijo, nOperacion
'        Case gCTSExtCancEfec
'            frmCapExtornos.Inicia gCTSCancEfec, sDescOperacion, gCapCTS, nOperacion
'        Case gCTSExtCancTransf
'            frmCapExtornos.Inicia gCTSCancTransf, sDescOperacion, gCapCTS, nOperacion
'        Case gCTSExtCancTransfAbCta 'RIRO20131212 ERS137
'            frmCapExtornos.Inicia gCTSCancTransfBco, sDescOperacion, gCapCTS, nOperacion
'        Case gAhoExtDctoEmiExt
'            frmCapExtornos.Inicia gAhoDctoEmiExt, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtDctoEmiOP
'            frmCapExtornos.Inicia gAhoDctoEmiOP, sDescOperacion, gCapAhorros, nOperacion
'        Case gAhoExtCargoComDivAho 'JUEZ 20130906
'            frmCapExtornos.Inicia gAhoCargoCobroComDiversasAho, sDescOperacion, gCapAhorros, nOperacion
'
'        Case gCMACOTAhoExtDepEfec
'            frmCapExtornos.Inicia gCMACOTAhoDepEfec, sDescOperacion, gCapAhorros, nOperacion
'        Case gCMACOTAhoExtDepChq
'            frmCapExtornos.Inicia gCMACOTAhoDepChq, sDescOperacion, gCapAhorros, nOperacion
'        Case gCMACOTAhoExtRetEfec
'            frmCapExtornos.Inicia gCMACOTAhoRetEfec, sDescOperacion, gCapAhorros, nOperacion
'        Case gCMACOTAhoExtRetOP
'            frmCapExtornos.Inicia gCMACOTAhoRetOP, sDescOperacion, gCapAhorros, nOperacion
'        Case "137000"
'            frmCapExtornos.Inicia "107001", sDescOperacion, gCapAhorros, nOperacion
'
'        Case gCredExtPago
'            frmCredExtornos.Caption = "Extorno Pago de Crédito"
'            frmCredExtornos.ExtornoPagos
'        Case gCredExtDesemb
'            frmCredExtornos.Caption = "Extorno Desembolso de Crédito"
'            frmCredExtornos.ExtornoDesemb
'        Case gCredExtCredRefina 'LUCV20160520, ERS004-2016
'            frmCredExtornos.Caption = "Extorno de Creditos Refinanciados"
'            frmCredExtornos.ExtornoVigencia 'End, LUCV
'        Case gCredExtPagoLote
'             frmCredExtornos.Caption = "Extorno Pago en Lote de Crédito"
'            frmCredExtPagoLote.Show 1
'        'WIOR 20131126 *************************************************
'        Case gCredExtPagoHonramiento
'            frmCredExtornos.Caption = "Extorno Pago Honramiento"
'            frmCredExtornos.ExtornoPagosHonramiento
'        'WIOR FIN ******************************************************
'        'FRHU 20150520 ERS022-2015 *************
'        Case gCredExtPagoTransfFocmacm
'            'frmColRecExtornoOpe.Inicio nOperacion, "Extorno de Pago de Credito Transferido FOCMACM"
'            frmColRecExtornoOpe.Inicio nOperacion, "Extorno de Pago de Crédito Transferido FOCMAC" 'FRHU 20150817 OBSERVACION
'        'FIN FRHU 20150520***
'        '***Agregado por ELRO el 20130717, según RFC1306270002****
'        Case gCapExtConSerPagDeb
'            frmCapServicioPagoDebitoExtorno.Show 1
'        '***Fin Agregado por ELRO el 20130717, según RFC1306270002
'        ' Agregado por RIRO el 20130401
'        Case gExtornoDepositoRecaudo
'            frmCapExtornoServicioRecaudo.Show 1
'        ' Fin RIRO
'        Case gExtornoServCobRegDebitoAuto 'JUEZ 20150130
'            frmServCobDebitoAutoExt.Show 1
'
'        'SEGURO DE TARJETAS DE DÉBITO
'        Case gCapExtCargoAfilSegTarjeta
'            frmSegTarjetaAfiliacionExt.Inicia nOperacion
'        'FRHU 20150128 ERS048-2014
'        Case gCapExtNotaDeCargo
'            frmCapExtornos.Inicia gCapNotaDeCargo, sDescOperacion, gCapAhorros, nOperacion
'        Case gCapExtNotaDeAbono
'            frmCapExtornos.Inicia gCapNotaDeAbono, sDescOperacion, gCapAhorros, nOperacion
'        'FIN FRHU 20150128
'        'MADM 20111214
'        Case 109006
'            frmcredExtornoPagoBN.Caption = "Extorno Pago BN - Convenio / Corresponsalia"
'            frmcredExtornoPagoBN.Show 1
'        'END MADM
'        'SERVICIOS
'        'COBRANZA de Servicios
'        Case gServCobSedalib, gServCobHidrandina, gServCobEdelnor
'            'frmCapServicios.inicia (nOperacion)
'        Case gServCobFideicomiso
'            'frmCapFideicomiso.Show 1
'        Case gServCobSATTInfraccion
'            'frmServCobranzaSat.Show 1
'        Case gServCobSATTReciboDerecho
'            'frmServCobranzatributos.Show 1
'        Case gServCobSATTReciboDerechoOficEsp
'            'frmServCobranzaTributosOE.Show 1
'        Case gServCobFoncodes
'            'frmCapFoncodes.Show 1
'        Case gServCobPlanBici
'            'frmCapPlanBici.Show 1
'        'MADM 20110321
'        Case gServCobServConv
'            'frmCredPagoServicios.Show 1
'        'END MADM
'        Case gServCobAfilSegTarj 'JUEZ 20150112
'            frmSegTarjetaAfiliacion.InicioOpc
'        Case gServCobDebitoAuto 'JUEZ 20150130
'            frmServCobDebitoAuto.Inicia gServCobDebitoAuto
'        Case gServCobDebitoAutoEdit 'JUEZ 20150130
'            frmServCobDebitoAuto.Inicia gServCobDebitoAutoEdit
'        Case gServCobSepelioPrima 'RECO20151124 ERS073-2015
'            frmSegSepelioCobroPrima.Show 1
'        Case gServActSepelioManual 'RECO20151124 ERS073-2015
'            frmSegSepelioAfiliacion.IniciaAfilManual
'        'RECO20160209 ERS073-2016
'        'EXTORNO SEGURO SEPELIO
'        Case 290005
'            frmCapExtornos.Inicia "200380", sDescOperacion, gCapAhorros, nOperacion
'        Case 290006
'            'frmOpEspecialesExt.Ini 300150, sDescOperacion
'            frmOpEspecialesExt.Ini 290006, sDescOperacion
'        Case 290007
'            frmOpEspecialesExt.Ini 290007, sDescOperacion
'        'RECO FIN
'
'        'EXTORNOS de Servicios
'        Case gServExtCobFideicomiso, gServExtCobHidrandina, _
'            gServExtCobSedalib, gServExtCobEdelnor, _
'            gServExtCobSATTInfraccion, gServExtCobSATTReciboDerecho, gServExtCobSATTReciboDerechoOficEsp
'            frmCajeroExtornos.Show 1
'
'        'Extornos de Giros
'        Case gServExtGiroApertEfec
'            frmCapExtornos.Inicia gServGiroApertEfec, sDescOperacion, gGiro, nOperacion
'        Case gServExtGiroCancEfec
'            frmCapExtornos.Inicia gServGiroCancEfec, sDescOperacion, gGiro, nOperacion
'
'        'Otras Operaciones
'        Case gOtrOpeDepCtaBcoEfec
'            frmOtrOpeDepCtaBco.Inicia nOperacion, sDescOperacion
'
'        'Regularizacion de Sobrante y Faltante
'        Case gOtrOpePagoFaltante
'            'MAVM 20120328 ***
'            'frmRegularizaSobFal.Ini sDescOperacion
'            frmRegularizaSobFal.Ini sDescOperacion, nOperacion
'            '***
'        'Ingresos
'        Case 300407
'            frmIngDevConv.Inicio nOperacion, "", "", sDescOperacion
'
'        'RIRO20150108 Ingresos por pago de cajero corresponsal
'        Case gOtrOpePagoRecaudoCajeroCorresponsal
'            Dim oPagoCC As New frmDepositoCajeroCorresponsal
'            oPagoCC.Inicio nOperacion
'            Set oPagoCC = Nothing
'        'END RIRO *********************************************
'
'        'Add By GITU 29-11-2012
'        Case 300470
'            frmComisionRepTarj.Show 1
'        'End GITU
'
'        Case 300401 To 300490
'            frmOpEspeciales.Ini nOperacion, sDescOperacion
'        'ARCV 13-03-2007
'        Case 300493 To 300498
'            frmOtrOpeDepCtaBco.Inicia nOperacion, sDescOperacion
'        '---------
'        'Case gOtrOpeTransBancos
'        '   frmOpEspeciales.Ini nOperacion, sDescOperacion
'        'Case gOtrOpeIngresosoCajaGeneral
'        '   frmOpEspeciales.Ini nOperacion, sDescOperacion
'        'Egresos
'        Case gOtrOpeAhoOtrosEgresos
'            frmOpEspeciales.Ini nOperacion, sDescOperacion
'        Case gOtrOpeDepositoBancos
'            frmOpEspeciales.Ini nOperacion, sDescOperacion
'        Case gOtrOpeDevolucionCredPersonal
'            'frmOpEspeciales.Ini nOperacion, sDescOperacion
'            frmCajeroOpeDevCredPers.Inicia nOperacion, sDescOperacion
'        'MADM 20111227 - 20110930
'        Case gOtrOpeEgresoComisionDepCtaRecaudadora 'EJVG20120417
'            frmOtrOpeDepCtaBco.Inicia nOperacion, sDescOperacion
'        Case gOtrOpeEgresoDevSobranteOtrasOpeChq 'EJVG20140408
'            frmOpeDevSobrante.Inicio nOperacion, sDescOperacion
'        Case gOtrOpeEgresoDevSobranteOtrasOpeVoucher 'RIRO20140530 ERS017
'            frmOpeDevSobranteVoucher.Inicio nOperacion, sDescOperacion
'        Case gOtrOpeCastigoDevolucionCredPersonal
'            'frmOpEspeciales.Ini nOperacion, sDescOperacion
'            frmCajeroOpeDevCredPers.Inicia nOperacion, sDescOperacion
'        'MADM 20120328 - COMENTADO X MADM 20120127
'        Case 300509
'            frmCajeroOpeDevSob.Inicia nOperacion, sDescOperacion
'        'END MADM
'        Case 300505
'            frmCajeroOpeEgreRef.Inicia nOperacion, sDescOperacion
'        '***Agregado por ELRO el 20120420, según OYP-RFC005-2012
'        Case gOtrOpeDesParGas
'            Call frmCapARendir.iniciarDesembolso(gOtrOpeDesParGas, "DESEMBOLSO PARA OTROS GASTOS")
'        Case gOtrOpeDesParVia
'            Call frmCapARendir.iniciarDesembolso(gOtrOpeDesParVia, "DESEMBOLSO PARA VIÁTICOS")
'        Case gOtrOpeDesParCaj
'            Call frmCapCajaChica.Show
'
'        '*** RIRO 20130702 SEGUN TI-ERS083-2013 ****************
'        Case gOtrOpeEgresosDirectos
'            frmOperEgresosEfectivo.Inicia
'        '*** FIN RIRO ******************************************
'
'        'RIRO20150527 ERS162-2014 ********
'        Case gotrOpeDepUtilidadesEfect
'            frmUtilidadesTramaPago.Show 1
'        'END RIRO ************************
'
'        '***Agregado por ELRO***********************************
'        Case 300504 To 300590
'            frmOpEspeciales.Ini nOperacion, sDescOperacion
'        'ARCV 13-03-2007
'        Case 300593 To 300599
'            frmOtrOpeDepCtaBco.Inicia nOperacion, sDescOperacion
'        '---------
'
'        'Formulario de Egresos de Desembolso de Caja Chica
'        Case gOtrOpeEgresosDesemCajaChica
'            FrmCajaGenDesemPendiente.Show 1
'        'Tarifas
'        Case gOtrOpeDuplicadoTarjeta
'            frmOpEspeciales.Ini nOperacion, sDescOperacion
'        Case gOtrOpeVentaListados
'            frmOpEspeciales.Ini nOperacion, sDescOperacion
'        Case gOtrOpeConsatnciaCancelacionCredito
'            frmOpEspeciales.Ini nOperacion, sDescOperacion
'        Case gOtrOpeElaboracionContrato
'            frmOpEspeciales.Ini nOperacion, sDescOperacion
'        Case "300605"
'            frmOpEspeciales.Ini nOperacion, sDescOperacion
'        Case 300606 To 300690
'            frmOpEspeciales.Ini nOperacion, sDescOperacion
'
'        Case gOtrOpeExtorno
'            frmOpEspecialesExt.Ini nOperacion, sDescOperacion
'
'        Case gCajaGenDesemPendienteExtorno
'            FrmCajaGenDesemPendienteExtorno.Show 1
'        Case gOtrOpePenalidadEcoTaxi 'EJVG20120622
'            frmPenalidadEcoTaxi.Show 1
'
'        'JUEZ 20130411 ***************************
'        Case gComisionReprogCredito
'            'frmOpeComisionReprogCred.Show 1 'JUEZ 20130528
'            frmOpeComisionReprogCred.Inicia gComisionReprogCredito, "Reprogramación de Crédito"
'        Case gComisionEvalPolEndosada
'            frmOpeComisionOtros.Inicia gComisionEvalPolEndosada, gColPParamComEvalPolEndosada, "Evaluación de Póliza Endosada", "Comision evaluación de póliza endosada", "EVALUACION  POLIZA ENDOSADA"
'        Case gComisionDupTasacion
'            frmOpeComisionOtros.Inicia gComisionDupTasacion, gColPParamComDupTasacion, "Duplicado de Tasación", "Comision duplicado de tasación", "DUPLICADO DE TASACION"
'        Case gComisionConsultaRENIEC
'            frmOpeComisionOtros.Inicia gComisionConsultaRENIEC, gColPParamComConsultaRENIEC, "Consulta RENIEC", "Comision consulta RENIEC", "CONSULTA RENIEC"
'        'END JUEZ ********************************
'        'JUEZ 20130528 ***************************
'        Case gComisionEnvioEstadoCta
'            frmOpeComisionReprogCred.Inicia gComisionEnvioEstadoCta, "Envío de estado de cuenta"
'        'END JUEZ ********************************
'        Case gComisionDiversasAho 'JUEZ 20130829
'            'frmOpeComisionDiversasAho.Inicia 'JUEZ 20150930
'
'        'COMISIONES
'        'JUEZ 20150928 *****************************
'        Case gComiAhoReposicionTarjeta
'            frmComisionRepTarj.Show 1
'        Case gComiAhoContMoneda
'            frmOpeComisionContMoneda.Show 1
'        Case gComiAhoDiversas
'            frmOpeComisionDiversasAho.Inicia "A"
'
'        Case gComiCredEmisionRenovCF
'            FrmCFComision.Show 1
'        Case gComiCredModifCF
'            frmCFComisionMod.Show 1
'        Case gComiCredReprogCred
'            frmOpeComisionReprogCred.Inicia nOperacion, sDescOperacion
'        Case gComiCredEvalPolizaEnd
'            frmOpeComisionOtros.Inicia nOperacion, gColPParamComEvalPolEndosada, sDescOperacion, "Comision " & sDescOperacion, "EVALUACION  POLIZA ENDOSADA"
'        Case gComiCredEstadoCta
'            frmOpeComisionReprogCred.Inicia nOperacion, sDescOperacion
'        Case gComiCredDupCronograma
'            frmOpeComisionCred.Inicia nOperacion, sDescOperacion, gColPParamComiDupCronograma
'        Case gComiCredConstNoAdeudo
'            frmOpeComisionOtros.Inicia nOperacion, gColPParamComiConstNoAdeudo, sDescOperacion, "Comision " & sDescOperacion, "CONSTANCIA DE NO ADEUDO"
'        Case gComiCredDiversas
'            frmOpeComisionDiversasAho.Inicia "C"
'        Case gComiOtrServBusqRegSUNARP
'            frmOpEspeciales.Ini gComiOtrServBusqRegSUNARP, sDescOperacion
'        'END JUEZ **********************************
'
'        'Operaciones de Credito
'        Case gCredDesembEfec
'            frmCredDesembAbonoCta.DesembolsoEfectivo gCredDesembEfec
'
'        Case gCredDesembCtaNueva
'            frmCredDesembAbonoCta.DesembolsoCargoCuenta gCredDesembCtaNueva
'        Case gCredDesembCheque
'            frmCredDesembAbonoCta.DesembolsoConCheque gCredDesembCheque
'
'        Case gCredPagNorNorEfec
'            frmCredPagoCuotas.Inicia gCredPagNorNorEfec   'Estaba Comentado
'        Case gCredPagAnticipAdelantoCuota 'JUEZ 20150415
'            frmCredPagoCuotasEspecial.Inicia gCredPagAnticipAdelantoCuota
'
'        Case gCredVigenciaRefina 'Agregado por ***LUCV 20160505- (ERS004-2016)
'            frmCredRefinancVigencia.RefinanciarCredito gCredVigenciaRefina '***Fin LUCV
'
'        'WIOR 20131126 ******************************************
'        Case gCredPagHonramiento
'            frmCredHonradosPago.Show 1
'        'WIOR FIN ***********************************************
'
'        'WIOR 20160425 ***
'        Case gCredPagLiqSegDes
'            frmGestionSiniestroLiquidar.Show 1
'        'WIOR FIN ********
'
'        'ALPA 20110819*****************************************
'        Case gCredPagLeasingCU
'            frmCredpagoCuotasLeasingDetalle.Inicia gCredPagLeasingCU
'        '******************************************************
'        'ALPA 20140621**************************************************************
'         Case gCredPagLeasingPC
'            Call frmCredpagoCuotasLeasingDetalle.Inicia(gCredPagLeasingPC, 2)
'        '***************************************************************************
'        Case gCredPagoCuotasEcotaxi 'EJVG20120518
'            'frmCredPagoCuotaEcotaxi.Inicio
'            Call frmCredPagoCuotaEcotaxi.Inicio(gCredPagoCuotasEcotaxi, gsOpeDesc) 'EJVG20130611
'        Case gCredPagoCuotasEcotaxiCoberturaOG 'EJVG20130611
'            Call frmCredPagoCuotaEcotaxi.Inicio(gCredPagoCuotasEcotaxiCoberturaOG, gsOpeDesc)
'        Case gCredPagNorRfaEfec
'            'FrmCreOpeRFA.Show vbModal
'        Case gCredPagNorNorDacion
'            'frmCredDacionPago.Show 1
'        Case gCredPagLote
'            frmCredPagoLote.Show 1
'     'madm 20100517
'        Case 102101
'            frmCredPagoConvenioBcoNac.Show
'        Case 102102
'            frmCredPagoConvenioBcoCred.Show
'        'FRHU 20150415 ERS022-2015
'        Case 103100
'            frmCredPagoTransferidos.Show 1
'        'FRHU FIN 20150415
'''        Case 901040
'''            frmCajeroCorte.Show
'         Case 901029
'            frmCajeroExtornos.Show 1
'    'end madm
'
'        'Operaciones de Pignoraticio
'
'        '***PEAC 20090504 Acta 022-2009 para diferenciar entre desembolso en efectivo y abono en cuenta
''        Case gColPOpeDesembolso
''            frmColPDesembolso.Show 1
'
'        Case gColPOpeDesembolsoEFE
'            frmColPDesemb.DesembolsoEfectivo gColPOpeDesembolsoEFE
'        Case gColPOpeDesembolsoAboCta
'            frmColPDesemb.DesembolsoCargoCuenta gColPOpeDesembolsoAboCta
'        Case 120205 'RECO20140129 ERS002 - RECO-N Se Agrego nueva opcion
'            frmColPDesembCredAmpliado.Show 1 'RECO20140129 ERS002 - RECO-N Se Agrego Nueva Opcion
'        '*** PEAC ***********************************
'        Case 120206 'RECO
'            frmColPDesemCampAdjudicado.Show 1 'reco
'        Case gColPOpeRenovacEFE
'            frmColPRenovacion.Inicio nOperacion, "Renovacion Pignoraticio", "", ""
'        Case gColPOpeRenovacCHQ
'            frmColPRenovacion.Inicio nOperacion, "Renovacion Pignoraticio", "", ""
'        Case gColPOpeCancelacEFE
'            frmColPCancelacion.Inicio nOperacion, "Cancelacion Pignoraticio", "", ""
'        Case gColPOpeCancelacCHQ
'            frmColPCancelacion.Inicio nOperacion, "Cancelacion Pignoraticio", "", ""
'        Case gColPOpeImpDuplicado
'            frmColPDuplicadoContrato.Show 1
'
'        Case gColPOpeAmortizEFE
'            '*** PEAC 20170127 - esta opcion será reemplazada por
'            '*** la operación 121400
'            'frmColPAmortizacion.Inicio nOperacion, "Amortizacion Pignoraticio", "", ""
'
'        Case gColPOpeDevJoyas
'            frmColPRescateJoyas.Show 1
'        'Add By GITU 10-07-2013
'
'        '*** PEAC 20161024
'        Case "121400"  ' PAGOS PARCIALES DE CRED PIGNORARTICIOS
'            frmColPPagosParciales.Inicio nOperacion, "Amortizacion Pignoraticio", "", ""
'
'        Case "122700"
'            frmColPRenovacion.Inicio nOperacion, "Renovacion Pignoraticio", "", ""
'        'End GITU
'        Case gColPOpeVtaSubasta
'            frmColPSubastaRegVenta.Inicio ("0000")
'        Case "122900"
'            frmColPRecuperacionReg.Inicio ("0000")
'            Set clscol = Nothing
'
'        Case gColPOpePagSobrante
'            frmColPPagoSobranteRemate.Show 1
'        Case gColPOpeVtaRemate
'            frmColPRemateRegVenta.Inicio ("")
'        Case gColPOpeVtaSubasta
'            frmColPSubastaRegVenta.Inicio ("")
'
'        '*** PEAC 20090313
'        Case gColPOpePagSobraAdjudicado  '  "122300"
'            frmColPPagoSobranteAdjudicacion.Show 1
'
'        'Case Duplicado sin Costo Chimbote
''        Case 121700
''            FrmColPDuplicadoContratoCostoCero.Show 1
''
'        Case gColPOpeCobCusDiferida
'            frmColPCustodiaDiferida.Show 1
'
'        'Extornos de Pignoraticio
'        Case gColPOpeExtDesemb
'            frmColPExtornoOpe.Inicio nOperacion, "Desembolso"
'        'RECO20140207 ERS002**************************
'        Case gColPOpeExtDesembAmpliado
'            frmColPExtornoOpe.Inicio nOperacion, "Desembolso por Ampliación"
'        'RECO FIN*************************************
'        Case "129401"
'            frmColPExtornoOpe.Inicio nOperacion, "Amortizacion"
'        Case "129700"
'            'frmColPExtornoOpe.Inicio nOperacion, "Amortizacion"
'        Case gColPOpeExtRenov
'            frmColPExtornoOpe.Inicio nOperacion, "Renovacion"
'        Case gColPOpeExtCance
'            frmColPExtornoOpe.Inicio nOperacion, "Cancelacion"
'        Case gColPOpeExtDupli
'            frmColPExtornoOpe.Inicio nOperacion, "Duplicado"
'        Case gColPOpeExtDevJoyas
'            frmColPExtornoOpe.Inicio nOperacion, "Devolucion Joyas"
'        Case "129300"
'            frmColPExtornoOpe.Inicio nOperacion, "Custodia Diferida"
'        Case "129701" ' Venta Remate
'            frmColPExtornoOpe.Inicio nOperacion, "Venta en Remate"
'        Case "129702" ' Pago Sobrante
'            frmColPExtornoOpe.Inicio nOperacion, "Pago de Sobrante"
'        Case "129703" ' Venta Adjudicado
'            frmColPExtornoOpe.Inicio nOperacion, "Venta de Adjudicado"
'        Case "129704" ' Venta Adjudicado
'            frmColPExtornoOpe.Inicio nOperacion, "Recuperacion de Adjudicado"
'
'        '*** PEAC 20090316
'        Case "129705" ' Sobrante de adjudicado
'            frmColPExtornoOpe.Inicio nOperacion, "Sobrante de Adjudicado"
'
'        Case "129801"
'            frmColPExtornoOpe.Inicio nOperacion, "Renovación"
'        Case "129802"
'            frmColPExtornoOpe.Inicio nOperacion, "Cancelación"
'        Case "129803"
'            frmColPExtornoOpe.Inicio nOperacion, "Amortización"
'
'        'Operaciones de Carta Fianza
'        Case gColCFOpeComisEfe
'            FrmCFComision.Show 1
'
'        'Extorno de Carta Fianza
'        Case gColCFOpeExtComis
'            frmCFExtornoOpe.Inicio nOperacion, "Comision Carta Fianza"
'        'WIOR 20120806********************
'        Case gColCFOpeComisMod
'            frmCFComisionMod.Show 1
'        'WIOR FIN ************************
'        'Pago de Credito en Recuperaciones
'        Case gColRecOpePagJudSDEfe
'            frmColRecPagoCredRecup.Inicio nOperacion, "Pago Credito en Recuperaciones efectivo", "", "", True
'        Case "130206" 'JACA 20110819-Pago Credito por Adjudicacion
'            frmColPagoCredAdjudicacion.Show 1
'        'Case gColRecOpePagJudSDChq
'            'frmColRecPagoCredRecup.Inicio nOperacion, "Pago Credito en Recuperaciones con cheque", "", "", True
'
'        'Extornos de Cred en Recuperaciones
'        Case gColRecOpeExtTransfRecup  ' Transferencia de Credito a Recuperaciones
'            frmColRecExtornoOpe.Inicio nOperacion, "Extorno de Transferencia a Recuperaciones"
'        Case gColRecOpeExtPagRecup  ' Pago de Credito en Recuperaciones
'            frmColRecExtornoOpe.Inicio nOperacion, "Extorno de Pago de Credito en Recuperaciones"
'
'        'Extornos de Pigno
'        Case gPigOpeExtDesembolso
'            'frmPigExtornoOpe.Inicio nOperacion, "Extorno de Desembolso", "", ""
'        Case gPigOpeExtAmortizEFE
'            'frmPigExtornoOpe.Inicio nOperacion, "Extorno de Amortizacion", "", ""
'        Case gPigOpeExtCancelacEFE
'            'frmPigExtornoOpe.Inicio nOperacion, "Extorno de Cancelacion", "", ""
'        Case gPigOpeExtReusoLinea
'            'frmPigExtornoOpe.Inicio nOperacion, "Extorno Uso de Línea", "", ""
'        Case gPigOpeExtCobCusDiferida
'            'frmPigExtornoOpe.Inicio nOperacion, "Extorno Cobro de Custodia Diferida", "", ""
'        Case gPigOpeExtImpDuplicado
'            'frmPigExtornoOpe.Inicio nOperacion, "Extorno Duplicado de Contrato", "", ""
'        Case gPigOpeAnulaVentaJoya
'            'frmPigAnularVentaJoyas.Show 1
'        Case gPigOpeExtPagoSobrante
'            'frmPigExtornoOpe.Inicio nOperacion, "Extorno Pago de Sobrante de Remate", "", ""
'        'Case gPigOpeExtRescateJoya
'            'frmPigExtornoOpe.Inicio nOperacion, "Extorno de Rescate de Joyas", "", ""
'        Case gCredPagNorRfaEfec
'            'FrmCreOpeRFA.Show vbModal
'        Case gColoOpeExRFA
'           'FrmExtornoRFA.Show vbModal
'        'Case 300508 'remesa con cheque
'           'FrmCajRemCheque.Show vbModal
'        Case 300700
'            'FrmCajRemCheque.Show vbModal
'        '**DAOR 20091116 Extorno operaciones intercajas********************
'        Case 159101 'Extorno pago de crédito intercmac
'            frmPITCapExtornos.Inicia 104001, sDescOperacion, 0, nOperacion
'        Case 279101 'Extorno retiro intercmac
'            frmPITCapExtornos.Inicia 261001, sDescOperacion, gCapAhorros, nOperacion
'        Case 279102 'Extorno depósito intercmac
'            frmPITCapExtornos.Inicia 261002, sDescOperacion, gCapAhorros, nOperacion
'
'        '******************************************************************
'        'ALPA 20100126*****************************************************
'        Case gAhoDepositoEnLoteEfec
'            frmCapDepositosEnLote.Inicia gCapAhorros, 200257, "DEPOSITO EN LOTE EN EFECTIVO"
'        Case gAhoDepositoEnLoteCheq
'            frmCapDepositosEnLote.Inicia gCapAhorros, 200258, "DEPOSITO EN LOTE CON CHEQUE"
'        Case gAhoDepositoEnLoteCarg
'            frmCapDepositoLoteCargo.Show 1
'        '******************************************************************
'        'RIRO ERS017
'        Case gAhoDepositoHaberesEnLoteEfec
'            frmCapDepositoLote.iniciarFormulario gCapAhorros, gAhoDepositoHaberesEnLoteEfec, "Deposito de Haberes en Lote Efectivo"
'        Case gAhoDepositoHaberesEnLoteTransf
'            frmCapDepositoLote.iniciarFormulario gCapAhorros, gAhoDepositoHaberesEnLoteTransf, "Deposito de Haberes en Lote Transferencia Banco"
'        'END RIRO
'        '***Agregado por ELRO el 20120412, según OYP-RFC005-2012
'        Case gOtrOpeExtDesParGas
'            frmCapExtARendir.iniciarExtornoDesembolso gOtrOpeDesParGas, "EXTORNO DESEMBOLSO PARA OTROS GASTOS", gOtrOpeExtDesParGas
'        Case gOtrOpeExtDesParVia
'            frmCapExtARendir.iniciarExtornoDesembolso gOtrOpeDesParVia, "EXTORNO DESEMBOLSO PARA VIÁTICOS", gOtrOpeExtDesParVia
'        Case gOtrOpeExtDesParCaj
'            frmCapExtARendir.iniciarExtornoDesembolso gOtrOpeDesParCaj, "EXTORNO DESEMBOLSO PARA CAJA CHICA", gOtrOpeExtDesParCaj
'        '***Fin Agregado por ELRO*******************************
'        Case gOtrOpeExtPenalidadEcoTaxi 'EJVG20120630
'            frmExtornoPenalidadEcotaxi.Show 1
'        '***Agregado por ELRO el 20130401, segoun TI-ERS011-2013****
'        Case gAhoExtMigracion
'            frmCapExtornos.Inicia gAhoMigracion, sDescOperacion, gCapAhorros, nOperacion
'
'        '***Fin Agregado por ELRO el 20130401, segoun TI-ERS011-2013
'
'        '*********** Agregado por RIRO el 20130314 ***********
'
'         Case gDepositoRecaudo
'                    frmPagoServRecaudo.Show 1
'
'         'Fin RIRO *******************************************
'         '***Agregado por ELRO el 20130712, segoun RFC1306270002****
'         Case gCapExtConSerPagDeb
'              frmCapExtornos.Inicia gCapConSerPagDeb, sDescOperacion, gCapAhorros, nOperacion
'        '***Fin Agregado por ELRO el 20130712, segoun RFC1306270002
'
'        '*******GIPO**************  ERS070-2016
'        Case "123100"
'            frmDevolverSobranteAdjudicado.Show 1
'    End Select
'End Sub

Private Sub cmdAceptar_Click()
    Dim nodOpe As Node
    Dim sDesc As String
    Dim oCajero As COMNCajaGeneral.NCOMCajero
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    
    'RIRO20131411 COMENTADO
    'If Not VerificaRFIII Then Exit Sub ' RIRO SEGUN TI-ERS108-2013

    Set oCajero = New COMNCajaGeneral.NCOMCajero
    If Not oCajero.YaRealizoCierreAgencia(gsCodAge, gdFecSis) Then
        'Set oCajero = Nothing '*** PEAC 20140318
        Set nodOpe = tvwOperacion.SelectedItem

        '*** PEAC 20140318 se incluyo en este procedimeinto para que valide tambien al ingresar cod de ope. /MIOL 20130308, SEGUN ERS008-2013 ***************************
        If oCajero.GetBuscarConfirmacionEfec(gsCodUser, Format(gdFecSis, "yyyyMMdd"), gsCodAge) Then
            gsOpeCodTemp = CLng(nodOpe.Tag)
            gsOpeCod = "901017H"
            frmCajeroExtornos.Inicia "CAJERO - CONFIRMACIÓN DE HABILITACIÓN"
            gsOpeCod = gsOpeCodTemp
        End If
        Set oCajero = Nothing
        'END PEAC/MIOL ***************************************************

        If Not nodOpe Is Nothing Then
            sDesc = Mid(nodOpe.Text, 8, Len(nodOpe.Text) - 7)
            'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
            Dim bPermitirEjecucionOperacion As Boolean
            Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
                
            bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, CLng(nodOpe.Tag), "0")
               
            If bPermitirEjecucionOperacion Then
                EjecutaOperacion CLng(nodOpe.Tag), sDesc
            Else
                Unload Me
                Unload MDISicmact
            End If
            'fin Comprobacion si es RFIII
        End If
        Set nodOpe = Nothing
    Else
        MsgBox "Ya se realizó el cierre de caja de la agencia. No es posible realizar transacciones", vbExclamation, "Aviso"
        Unload Me
        Set oCajero = Nothing
    End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

'*** MAVM: Modulo de Auditoria 21/08/2008
' Para Mostrar seleccionado El Reporte de los Movimientos de las Cuentas de Ahorro
' por defecto en el Modulo de Auditoria
Public Sub Inicializar_MovimientosCtasAhorros(ByVal sCaption As String, ByVal rsUsu As ADODB.Recordset)
    'tvwOperacion.Nodes.Clear
    'LlenaArbol rsUsu

    tvwOperacion.Nodes(Index200000).Selected = True
    tvwOperacion_DblClick
    tvwOperacion.Nodes(Index200000).Expanded = True
    
    tvwOperacion.Nodes(Index201000).Selected = True
    tvwOperacion_DblClick
    tvwOperacion.Nodes(Index201000).Expanded = True

    tvwOperacion.Nodes(Index201002).Selected = True
    tvwOperacion_DblClick
    tvwOperacion.Nodes(Index201002).Expanded = True

    tvwOperacion.Enabled = False
    tvwOperacion.HideSelection = False
End Sub
'*** MAVM: Modulo de Auditoria 21/08/2008

Private Sub tvwOperacion_DblClick()
Dim nodOpe As Node
Dim gsOpeCodTemp As String 'MIOL 20130308, SEGUN ERS008-2013
Dim sDesc As String
Dim oCajero As COMNCajaGeneral.NCOMCajero

    'RIRO20131411 COMENTADO
    'If Not VerificaRFIII Then Exit Sub ' RIRO SEGUN TI-ERS108-2013

    If tvwOperacion.SelectedItem.Child Is Nothing Then
        Set oCajero = New COMNCajaGeneral.NCOMCajero

        If Not oCajero.YaRealizoCierreAgencia(gsCodAge, gdFecSis) Then
            Set nodOpe = tvwOperacion.SelectedItem 'MIOL 20130308, SEGUN ERS008-2013 Comente
            'MIOL 20130308, SEGUN ERS008-2013 ***************************
            If oCajero.GetBuscarConfirmacionEfec(gsCodUser, Format(gdFecSis, "yyyyMMdd"), gsCodAge) Then
                gsOpeCodTemp = CLng(nodOpe.Tag)
                gsOpeCod = "901017H"
                frmCajeroExtornos.Inicia "CAJERO - CONFIRMACIÓN DE HABILITACIÓN"
                gsOpeCod = gsOpeCodTemp
            End If
            'END MIOL ***************************************************
            Set oCajero = Nothing
            
            '***Agregado por ELRO el 20120312 según Acta N° 045-2012/TI-D
            If gsCodAge = "01" Or gsCodAge = "04" Or _
                gsCodAge = "09" Or gsCodAge = "31" Then
                Dim lsFechaValidador As String
                lsFechaValidador = validarFechaSistema
                If lsFechaValidador <> "" Then
                    If gdFecSis <> CDate(lsFechaValidador) Then
                        MsgBox "La Fecha de tu sesión en el Negocio no coincide con la fecha del Sistema", vbCritical, "Aviso"
                        Call SalirSICMACMNegocio
                        Unload Me
                        End
                    End If
                End If
                lsFechaValidador = ""
            End If
            '***Fin Agregado por ELRO************************************
            
            'Set nodOpe = tvwOperacion.SelectedItem 'MIOL 20130308, SEGUN ERS008-2013
                        
            If Not nodOpe Is Nothing Then
                sDesc = Mid(nodOpe.Text, 8, Len(nodOpe.Text) - 7)
                'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
                Dim bPermitirEjecucionOperacion As Boolean
                Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
                
                bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, CLng(nodOpe.Tag), "0")
               
                If bPermitirEjecucionOperacion Then
                     EjecutaOperacion CLng(nodOpe.Tag), sDesc
                Else
                    Unload Me
                    Unload MDISicmact
                End If
                'fin Comprobacion si es RFIII
            End If
            Set nodOpe = Nothing
        Else
            MsgBox "Ya se realizó el cierre de caja de la agencia. No es posible realizar transacciones", vbExclamation, "Aviso"
            Unload Me
            Set oCajero = Nothing
        End If
        

        
    End If

End Sub

Private Sub tvwOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim nodOpe As Node
    Dim sDesc As String
    Dim oCajero As COMNCajaGeneral.NCOMCajero
    
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    If Not oCajero.YaRealizoCierreAgencia(gsCodAge, gdFecSis) Then
        Set oCajero = Nothing
        
        '***Agregado por ELRO el 20120312 según Acta N° 045-2012/TI-D
        If gsCodAge = "01" Or gsCodAge = "04" Or _
            gsCodAge = "09" Or gsCodAge = "31" Then
            Dim lsFechaValidador As String
            lsFechaValidador = validarFechaSistema
            If lsFechaValidador <> "" Then
                If gdFecSis <> CDate(lsFechaValidador) Then
                    MsgBox "La Fecha de tu sesión en el Negocio no coincide con la fecha del Sistema", vbCritical, "Aviso"
                    Call SalirSICMACMNegocio
                    Unload Me
                    End
                End If
            End If
            lsFechaValidador = ""
        End If
        '***Fin Agregado por ELRO************************************
        
        Set nodOpe = tvwOperacion.SelectedItem
        If Not nodOpe Is Nothing Then
            sDesc = Mid(nodOpe.Text, 8, Len(nodOpe.Text) - 7)
            EjecutaOperacion CLng(nodOpe.Tag), sDesc
        End If
        Set nodOpe = Nothing
    Else
        MsgBox "Ya se realizó el cierre de caja de la agencia. No es posible realizar transacciones", vbExclamation, "Aviso"
        Unload Me
        Set oCajero = Nothing
    End If

End If
End Sub

Private Sub txtOperacion_KeyUp(KeyCode As Integer, Shift As Integer)
Dim nodOpe As Node
For Each nodOpe In tvwOperacion.Nodes
    If nodOpe.Tag = Trim(txtOperacion) Then
        tvwOperacion.SelectedItem = nodOpe
        Exit For
    End If
Next
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub
'***Agregado por ELRO el 20120312 según Acta N° 045-2012/TI-D
Private Sub SalirSICMACMNegocio()
Dim oSeguridad As New COMManejador.Pista
    Call oSeguridad.InsertarPista(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, "Salida del SICMACM Negocio" & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-20120118")
     If oSeguridad.ValidaAccesoPistaRF(gsCodUser) Then
            'ALPA 20160215**********************************
            Call oSeguridad.InsertarPistaSesion(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, 0)
            '***********************************************
            '0-Operaciones
            '1-Negocio
            
     End If
    Set oSeguridad = Nothing
End Sub

' *** RIRO SEGUN TI-ERS108-2013 ***
Private Function VerificaRFIII() As Boolean
    Dim rsRF3 As New ADODB.Recordset
    Dim sMensaje As String
    Set rsRF3 = ValidarRFIII
    sMensaje = ""
    If Not (rsRF3.EOF Or rsRF3.BOF) And rsRF3.RecordCount > 0 Then
        If gsCodCargo = "006005" Then    ' *** SI ES "SUPERVISOR"
            If Not rsRF3!bOpcionesSimultaneas And rsRF3!bModoSupervisor Then
                sMensaje = "Actualmente no puede realizar ninguna operación debido a que el RFIII se encuentra activo en modo supervisor " & vbNewLine & _
                           "por favor active el segundo perfil del RFIII para poder acceder a esta opción "
            End If
        ElseIf gsCodCargo = "007026" Then ' *** SI ES "RFIII"
            If rsRF3!bPerfilCambiado Then
                If rsRF3!bModoSupervisor Then
                    sMensaje = "Su perfil ha cambiado a modo supervisor, debe volver a acceder al sistema"
                Else
                    sMensaje = "Su perfil ha cambiado a modo normal, debe volver a acceder al sistema"
                End If
            End If
        End If
    End If
    If Len(sMensaje) > 0 Then
        MsgBox sMensaje, vbExclamation, "Aviso"
        If gsCodCargo = "006005" Then ' SUPERVISOR
            VerificaRFIII = False
            Exit Function
        ElseIf gsCodCargo = "007026" Then ' RFIII
            VerificaRFIII = False
            End
        End If
    End If
    VerificaRFIII = True
End Function
' *** FIN RIRO ***
'RECO20151111 ERS061-2015 *************************
Private Function VerificaGrupoPermisoPostCierre() As Boolean
    Dim oCons As New COMDConstSistema.DCOMGeneral
    Dim sGrupoAutorizado As String
    Dim nGrupoTmp1 As String
    Dim nGrupoTmp2 As String
    Dim i As Integer
    Dim j As Integer
            
    sGrupoAutorizado = oCons.LeeConstSistema(516)
    VerificaGrupoPermisoPostCierre = False
    For i = 1 To Len(sGrupoAutorizado)
        If Not Mid(sGrupoAutorizado, i, 1) = "," Then
            nGrupoTmp1 = nGrupoTmp1 & Mid(sGrupoAutorizado, i, 1)
        Else
            For j = 1 To Len(gsGruposUser)
                If Not Mid(gsGruposUser, j, 1) = "," Then
                    nGrupoTmp2 = nGrupoTmp2 & Mid(gsGruposUser, j, 1)
                Else
                    If nGrupoTmp1 = nGrupoTmp2 Then
                        VerificaGrupoPermisoPostCierre = True
                        Exit Function
                    End If
                    nGrupoTmp2 = ""
                End If
            Next
            nGrupoTmp1 = ""
        End If
    Next
End Function
'RECO FIN *****************************************
