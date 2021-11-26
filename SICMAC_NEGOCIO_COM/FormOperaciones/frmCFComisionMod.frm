VERSION 5.00
Begin VB.Form frmCFComisionMod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta Fianza - Comisión por Modificación"
   ClientHeight    =   4905
   ClientLeft      =   2265
   ClientTop       =   1845
   ClientWidth     =   7590
   Icon            =   "frmCFComisionMod.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   7335
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   300
         TabIndex        =   29
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5880
         TabIndex        =   28
         Top             =   195
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1620
         TabIndex        =   27
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Afianzado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   120
      TabIndex        =   17
      Top             =   660
      Width           =   7425
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Afianzado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   19
         Tag             =   "txtcodigo"
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2340
         TabIndex        =   18
         Tag             =   "txtnombre"
         Top             =   210
         Width           =   4920
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Acreedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   675
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   7410
      Begin VB.Label lblNomAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2340
         TabIndex        =   16
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   4890
      End
      Begin VB.Label lblCodAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   15
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Acreedor "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame fraDatos 
      Enabled         =   0   'False
      Height          =   660
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   7380
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5715
         MaxLength       =   255
         TabIndex        =   33
         Top             =   180
         Width           =   1290
      End
      Begin VB.TextBox TxtITF 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3270
         MaxLength       =   255
         TabIndex        =   31
         Top             =   180
         Width           =   1290
      End
      Begin VB.TextBox TxtComision 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1095
         MaxLength       =   255
         TabIndex        =   2
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5190
         TabIndex        =   34
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ITF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2895
         TabIndex        =   32
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label10 
         Caption         =   "Comisión "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.CommandButton CmdExaminar 
      Caption         =   "E&xaminar..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   180
      Width           =   1230
   End
   Begin VB.Frame FraCredito 
      Caption         =   "Carta Fianza"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   7395
      Begin VB.Label Label9 
         Caption         =   "Analista"
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   900
         Width           =   720
      End
      Begin VB.Label lblAnalista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   24
         Top             =   840
         Width           =   3420
      End
      Begin VB.Label lblFecVencCF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5700
         TabIndex        =   22
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento"
         Height          =   195
         Left            =   4740
         TabIndex        =   21
         Top             =   900
         Width           =   870
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Monto"
         Height          =   255
         Left            =   4740
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   4740
         TabIndex        =   8
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Modalidad"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblTipoCF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   6
         Top             =   240
         Width           =   3420
      End
      Begin VB.Label lblMontoApr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5700
         TabIndex        =   5
         Top             =   540
         Width           =   1590
      End
      Begin VB.Label lblModalidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   4
         Top             =   540
         Width           =   3420
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5700
         TabIndex        =   3
         Top             =   240
         Width           =   1590
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   390
      Left            =   120
      TabIndex        =   30
      Top             =   180
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   688
      Texto           =   "Cta Fianza"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   23
      Top             =   240
      Width           =   1605
   End
End
Attribute VB_Name = "frmCFComisionMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*  APLICACION : Carta Fianza
'*  ARCHIVO : frmCFComisionMod
'*  CREACION: 28/08/2012      AUTOR :WIOR
'*  MODIFICACION
'***************************************************************************
'*  RESUMEN:  Permite registrar el cobro de comision por modificación de la Carta Fianza
'***************************************************************************
Option Explicit

Dim vCodCta As String
Dim fpComision As Double
Dim fbComisionTrimestral  As Boolean

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    vCodCta = ActXCodCta.NroCuenta
    If KeyAscii = 13 Then
    Dim oCreditos As COMNCartaFianza.NCOMCartaFianzaValida
    Dim rsCreditos As ADODB.Recordset
    Set oCreditos = New COMNCartaFianza.NCOMCartaFianzaValida
    Set rsCreditos = oCreditos.RecuperaPersonasEstadoCF(Array(gColocEstVigNorm), , gdFecSis, 2, vCodCta)
        If Len(vCodCta) > 0 Then
            If rsCreditos.RecordCount > 0 Then
                Call CargaDatosR(vCodCta)
                ActXCodCta.Enabled = False
            Else
                MsgBox "Carta Fianza no cuenta con las especificaciones establecidas.", vbInformation, "Aviso"
                Call LimpiarControles
                ActXCodCta.SetFocus
                Set oCreditos = Nothing
                Set rsCreditos = Nothing
            End If
        Else
            Call LimpiarControles
            ActXCodCta.SetFocus
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    LimpiarControles
    ActXCodCta.SetFocus
End Sub

Private Sub cmdExaminar_Click()
'Solo cobra a CF Aprobadas
Dim lsCta As String
    lsCta = frmCFPersEstado.inicio(Array(gColocEstVigNorm), "Comision por Modificación de Carta Fianza", Array(gColCFComercial, gColCFPYME, gColCFTpoProducto), 2)
    If Len(Trim(lsCta)) > 0 Then
        ActXCodCta.NroCuenta = lsCta
        Call CargaDatosR(lsCta)
    Else
        Call LimpiarControles
    End If
End Sub

'PROCEDIMIENTO QUE CARGA LOS DATOS QUE SE REQUIEREN PARA EL FORMULARIO
Sub CargaDatosR(ByVal psCta As String)
Dim oCF As COMDCartaFianza.DCOMCartaFianza
Dim R As ADODB.Recordset
Dim loCFCalculo As COMNCartaFianza.NCOMCartaFianzaCalculos
Dim loConstante As COMDConstantes.DCOMConstantes
Dim loCFValida As COMNCartaFianza.NCOMCartaFianzaValida
Dim lbTienePermiso As Boolean
Dim lnComisionPagada As Double
Dim lnComisionCalculada As Double
Dim ldFechaAsi As Date
'----- MADM
Dim lafirma As frmPersonaFirma
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim Rf As ADODB.Recordset
'----- MADM

ActXCodCta.Enabled = False
    gITF.fgITFParamAsume (Mid(psCta, 4, 2))
    Set oCF = New COMDCartaFianza.DCOMCartaFianza
        Set R = oCF.RecuperaCartaFianzaComision(psCta)
    Set oCF = Nothing
    If R Is Nothing Then
        MsgBox "No se encuentra información de la Carta Fianza", vbInformation, "Aviso"
        Exit Sub
    End If
    If Not R.BOF And Not R.EOF Then
        lblcodigo.Caption = R!cPersCod
        lblNombre.Caption = PstaNombre(R!cPersNombre)
        lblCodAcreedor.Caption = R!cPersAcreedor
        lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
        lblEstado.Caption = fgEstadoColCFDesc(R!nPrdEstado)
        ldFechaAsi = R!dAsignacion
        
            lblTipoCF = IIf(IsNull(R!cConsDescripcion), "", R!cConsDescripcion)
            LblMoneda = IIf(Mid(Trim(psCta), 9, 1) = "1", "SOLES", "DOLARES")

            LblMoneda = IIf(Mid(Trim(psCta), 9, 1) = "1", "SOLES", "DOLARES")

        lblAnalista.Caption = PstaNombre(IIf(IsNull(R!cAnalista), "", R!cAnalista))
        lblMontoApr = IIf(IsNull(R!nMontoApr), "", Format(R!nMontoApr, "#0.00"))
        lblFecVencCF = IIf(IsNull(R!dVencApr), "", Format(R!dVencApr, "dd/mm/yyyy"))
    
        Set loConstante = New COMDConstantes.DCOMConstantes
            lblModalidad = loConstante.DameDescripcionConstante(gColCFModalidad, R!nModalidad)
        Set loConstante = Nothing
        
        Set loCFValida = New COMNCartaFianza.NCOMCartaFianzaValida

            lnComisionPagada = loCFValida.nCFPagoComision(psCta, gdFecSis, 1)

        Set loCFValida = Nothing
        Set loCFCalculo = New COMNCartaFianza.NCOMCartaFianzaCalculos
            If fbComisionTrimestral = False Then
                lnComisionCalculada = Format(loCFCalculo.nCalculaComisionCF(R!nMontoApr, DateDiff("d", ldFechaAsi, R!dVencApr), fpComision, CInt(Mid(Trim(psCta), 9, 1))), "#,##0.00")
            Else
                lnComisionCalculada = Format(loCFCalculo.nCalculaComisionTrimestralCF(R!nMontoApr, DateDiff("d", ldFechaAsi, R!dVencApr), R!nModalidad, Mid(R!cCtaCod, 9, 1), psCta, 6), "#,###0.00")
            End If
        Set loCFCalculo = Nothing
        If lnComisionPagada >= lnComisionCalculada Then
            MsgBox " Ya se realizo el Pago de Comisión para la Modificación de Carta Fianza", vbInformation, "Aviso"
            LimpiarControles
            Exit Sub
        Else
            txtComision.Text = Format(lnComisionCalculada - lnComisionPagada, "#,#0.00")
        End If
        
        If R!dVencApr < gdFecSis Then
            MsgBox "Fecha de Vencimiento es anterior a la Fecha Actual ", vbInformation, "Aviso"
        
        '************ firma
         If lblcodigo.Caption <> "" Then
             Set lafirma = New frmPersonaFirma
             Set ClsPersona = New COMDPersona.DCOMPersonas
            
             Set Rf = ClsPersona.BuscaCliente(lblcodigo.Caption, BusquedaCodigo)
             
             If Rf!nPersPersoneria = 1 Then
             Call frmPersonaFirma.inicio(Trim(lblcodigo.Caption), Mid(lblcodigo.Caption, 4, 2), False, True)
             End If
             Set Rf = Nothing
        End If
        '************ firma
        Exit Sub
        End If
        
        'color a txtComision a pagar
        If Mid(psCta, 6, 1) = "2" Then
            txtComision.BackColor = RGB(183, 232, 189)
        Else
            txtComision.BackColor = vbWhite
        End If
        
        '---- ITF
        If gITF.gbITFAplica = True Then
           If Not gITF.gbITFAsumidocreditos Then
                Me.TxtITF = Format(gITF.fgITFCalculaImpuesto(CDbl(Me.txtComision)), "#0.00")
                Me.txtTotal = Format(CDbl(Me.txtComision) + CDbl(Me.TxtITF), "#0.00")
            Else
                Me.TxtITF = Format(gITF.fgITFCalculaImpuesto(CDbl(Me.txtComision)), "#0.00")
                Me.txtTotal = Me.txtComision
            End If
        Else
            Me.TxtITF = Format(0, "#0.00")
            Me.txtTotal = Me.txtComision
        End If
        '---- ITF
        
         '************ firma
         If lblcodigo.Caption <> "" Then
             Set lafirma = New frmPersonaFirma
             Set ClsPersona = New COMDPersona.DCOMPersonas
            
             Set R = ClsPersona.BuscaCliente(frmCFPersEstado.vcodper, BusquedaCodigo)
             If Not R.BOF And Not R.EOF Then
                    If R!nPersPersoneria = 1 Then
                    Call frmPersonaFirma.inicio(Trim(frmCFPersEstado.vcodper), Mid(frmCFPersEstado.vcodper, 4, 2), False)
                    End If
             End If
             Set R = Nothing
        End If
        '************
        
        fraDatos.Enabled = True
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus

    End If
    
End Sub

Private Sub cmdGrabar_Click()
'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
'fin Comprobacion si es RFIII
    
Dim loNCartaFianza As COMNCartaFianza.NCOMCartaFianza
Dim loImprime As COMNCartaFianza.NCOMCartaFianzaImpre
Dim loPrevio As previo.clsprevio
Dim lsCadImprimir As String

Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnUltTrans As Integer
Dim lnMonto As Currency
vCodCta = ActXCodCta.NroCuenta
    
lnMonto = CCur(Me.txtComision.Text)



If ValidaDatos = False Then
    Exit Sub
End If

If MsgBox("Desea Grabar Cobranza de Comision Por Modificación de Carta Fianza", vbInformation + vbYesNo, "Aviso") = vbYes Then

    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
    Set loNCartaFianza = New COMNCartaFianza.NCOMCartaFianza
        
        Call loNCartaFianza.nCFComision(vCodCta, lsFechaHoraGrab, lsMovNro, lnMonto, gITF.gbITFAplica, gITF.gbITFAsumidocreditos, CCur(Val(Me.TxtITF)), 1)
    
    Set loNCartaFianza = Nothing
    
    ' *** Impresion
        Set loImprime = New COMNCartaFianza.NCOMCartaFianzaImpre
            lsCadImprimir = loImprime.nPrintReciboCFComision(gsNomAge, lsFechaHoraGrab, vCodCta, lblNombre, _
                    lblNomAcreedor, Me.lblMontoApr, lblFecVencCF, lnMonto, gsCodUser, "", gsCodCMAC, CDbl(Val(TxtITF.Text)), gImpresora, gbImpTMU, 1)
        Set loImprime = Nothing
        Set loPrevio = New previo.clsprevio
            loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
            Do While True
                If MsgBox("Desea ReImprimir Recibo ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    loPrevio.PrintSpool sLpt, lsCadImprimir, False, 22
                Else
                    Set loPrevio = Nothing
                    Exit Do
                End If
            Loop
         Set loPrevio = Nothing
    LimpiarControles
    cmdGrabar.Enabled = False
    'INICIO JHCU ENCUESTA 16-10-2019
    Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
    'FIN
End If

End Sub

Function ValidaDatos() As Boolean

'valida Monto de la Comision
If CDbl(Me.txtComision) <= 0 Then
    MsgBox "Monto de la Comision debe ser mayor a Cero", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

ValidaDatos = True
End Function

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColCFComercial, False)
        If sCuenta <> "" Then
            ActXCodCta.NroCuenta = sCuenta
            ActXCodCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    LimpiarControles
    CargaParametros
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Sub LimpiarControles()
   ActXCodCta.Enabled = True
   ActXCodCta.NroCuenta = fgIniciaAxCuentaCF
   lblcodigo.Caption = ""
   lblNombre.Caption = ""
   lblCodAcreedor.Caption = ""
   lblNomAcreedor.Caption = ""
   lblTipoCF.Caption = ""
   LblMoneda.Caption = ""
   lblMontoApr.Caption = ""
   lblModalidad.Caption = ""
   lblAnalista.Caption = ""
   lblFecVencCF.Caption = ""
   lblEstado.Caption = ""
   txtComision = ""
   Me.TxtITF = "0.00"
   Me.txtTotal = "0.00"
   fraDatos.Enabled = False
   cmdGrabar.Enabled = False
End Sub


'Carga los Parametros
Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Dim lcCons As COMDConstSistema.DCOMConstSistema

Dim lr As New ADODB.Recordset
Set loParam = New COMDColocPig.DCOMColPCalculos
    fpComision = loParam.dObtieneColocParametro(4001)
Set loParam = Nothing

Set lcCons = New COMDConstSistema.DCOMConstSistema
    Set lr = lcCons.ObtenerVarSistema()
        fbComisionTrimestral = IIf(lr!nConsSisValor = 2, True, False)
    Set lr = Nothing
Set lcCons = Nothing
End Sub



