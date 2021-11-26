VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCFRenovacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Renovación"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   Icon            =   "FrmCFRenovacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Aval"
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
      TabIndex        =   39
      Top             =   1880
      Width           =   7410
      Begin VB.Label lblNomAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   42
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   5010
      End
      Begin VB.Label lblCodAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   41
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Aval"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.Frame FraCredito 
      Caption         =   "Carta Fianza"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   7395
      Begin VB.TextBox TxtMonApr 
         Height          =   285
         Left            =   5760
         TabIndex        =   38
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Comision"
         Height          =   255
         Left            =   4680
         TabIndex        =   35
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblcomision 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5760
         TabIndex        =   34
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5760
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblModalidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   22
         Top             =   540
         Width           =   3420
      End
      Begin VB.Label lblTipoCF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   21
         Top             =   240
         Width           =   3420
      End
      Begin VB.Label Label8 
         Caption         =   "Modalidad"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   4680
         TabIndex        =   19
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Monto"
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento"
         Height          =   195
         Left            =   4680
         TabIndex        =   16
         Top             =   1320
         Width           =   870
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
         Left            =   5760
         TabIndex        =   15
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblAnalista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   14
         Top             =   840
         Width           =   3420
      End
      Begin VB.Label Label9 
         Caption         =   "Analista"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   900
         Width           =   720
      End
   End
   Begin VB.Frame fraDatos 
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
      Height          =   705
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   7380
      Begin VB.TextBox TxtPeriodo 
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
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   27
         Top             =   240
         Width           =   660
      End
      Begin MSMask.MaskEdBox txtFecVencNueva 
         Height          =   315
         Left            =   6000
         TabIndex        =   28
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtFecEmiNue 
         Height          =   315
         Left            =   1440
         TabIndex        =   33
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nueva Emision"
         ForeColor       =   &H80000006&
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   31
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Vencimiento:"
         ForeColor       =   &H80000006&
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   30
         Top             =   240
         Width           =   1440
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
      TabIndex        =   7
      Top             =   1200
      Width           =   7410
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Acreedor "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblCodAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   9
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label lblNomAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   8
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   5010
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
      TabIndex        =   3
      Top             =   540
      Width           =   7425
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   6
         Tag             =   "txtnombre"
         Top             =   210
         Width           =   5025
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   5
         Tag             =   "txtcodigo"
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Afianzado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   7380
      Begin VB.CommandButton cmdCheckListCFRenv 
         Caption         =   "CheckList"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   43
         ToolTipText     =   "CheckList"
         Top             =   195
         Width           =   1005
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1680
         TabIndex        =   25
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5940
         TabIndex        =   2
         Top             =   195
         Width           =   1155
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   180
         Width           =   1155
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   390
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   3645
      _extentx        =   6429
      _extenty        =   688
      texto           =   "Cta Fianza"
      enabledcmac     =   -1
      enabledcta      =   -1
      enabledprod     =   -1
      enabledage      =   -1
   End
   Begin VB.Label lblPoliza 
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
      Left            =   6000
      TabIndex        =   37
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Num Folio :"
      Height          =   195
      Index           =   1
      Left            =   6000
      TabIndex        =   36
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Renovación: "
      Height          =   195
      Index           =   0
      Left            =   4440
      TabIndex        =   29
      Top             =   0
      Width           =   960
   End
   Begin VB.Label LblRenovacion 
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
      Left            =   4440
      TabIndex        =   26
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmCFRenovacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*  APLICACION : Carta Fianza
'*  ARCHIVO : frmCFRechazar
'*  CREACION: 10/09/2002     - AVMM
'*************************************************************************
'*  RESUMEN: RENOVACION DE CARTA FIANZA
'***************************************************************************

Option Explicit
Dim vCodCta As String
Dim fbComisionTrimestral As Boolean
Dim loCFCalculo As COMNCartaFianza.NCOMCartaFianzaCalculos
Dim lcCons As COMDConstSistema.DCOMConstSistema
Dim loParam As COMDColocPig.DCOMColPCalculos
Dim fpComision As Double
Dim lr As New ADODB.Recordset
Dim sCodCta As String
Dim lnModalidad As Integer
Dim objPista As COMManejador.Pista 'MAVM 20100625 BAS II
'WIOR 20130313 *******************************************
Dim fdVigCol As Date
Dim fdEmisionAnt As Date
Dim fdVencAnt As Date
Dim fnMontoRAnt As Double
Dim fnMotivo As Integer
Dim fdPrdEstado As Date
Dim fnPrdEstado As Integer
Dim oCFRenNew As COMDCartaFianza.DCOMCartaFianza
Dim rsCFRenNew As ADODB.Recordset
'WIOR FIN ************************************************
Dim bCheckList As Boolean 'JOEP20190124 CP

Private Sub cmdCancelar_Click()
    LimpiarControles
    ActXCodCta.SetFocus
    bCheckList = False 'JOEP20190124 CP
End Sub

'PROCEDIMIENTO QUE CARGA LOS DATOS QUE SE REQUIEREN PARA EL FORMULARIO
Sub CargaDatosR(ByVal psCodCta As String)
    Dim oCF As COMDCartaFianza.DCOMCartaFianza 'DCartaFianza
    Dim loConstante As COMDConstantes.DCOMConstantes 'DConstante
    Dim R As New ADODB.Recordset
    Dim dFecha As Date
    
    Dim lbTienePermiso As Boolean
    'WIOR 20130313 *******************************************
    Dim ldNEmision As Date
    Dim ldNVenc As Date
    Dim lnMontoR As Double
    Dim lnPeriodo As Integer
    'WIOR FIN ************************************************
    Dim oDCOMCredito As New COMDCredito.DCOMCredito 'LUCV20171212, Agregó según observación SBS
    
    'On Error GoTo ErrorCargaDat
    dFecha = gdFecSis
    ActXCodCta.Enabled = False
    bCheckList = False 'JOEP20190124 CP
    Set oCF = New COMDCartaFianza.DCOMCartaFianza
        Set R = oCF.RecuperaCartaFianzaRenovacion(psCodCta, dFecha)
    Set oCF = Nothing
    
    If Not R.BOF And Not R.EOF Then
        'WIOR 20130313 *******************************************
        Set oCFRenNew = New COMDCartaFianza.DCOMCartaFianza
        Set rsCFRenNew = oCFRenNew.ObternerAutRenovacionNro(psCodCta, IIf(IsNull(R!nRenovacion), "0", R!nRenovacion), "0,1")
        If rsCFRenNew.RecordCount > 0 Then
            If CInt(rsCFRenNew!nEstado) = 1 Then
                ldNEmision = Format(rsCFRenNew!dNEmision, "dd/mm/yyyy")
                ldNVenc = Format(rsCFRenNew!dNVencimiento, "dd/mm/yyyy")
                lnMontoR = rsCFRenNew!nMonto
                lnPeriodo = rsCFRenNew!nPeriodo
            Else
                lnMontoR = CDbl(R!nMontoApr)
                MsgBox "Aun no Pago la comisión", vbInformation, "Aviso"
                Exit Sub
            End If
        Else
            lnMontoR = CDbl(R!nMontoApr)
            MsgBox "No se realizo el proceso correctamente", vbInformation, "Aviso"
            Exit Sub
        End If
        
        '***** LUCV20171212, Agregío según observación SBS *****
        If oDCOMCredito.verificarExisteAutorizaciones(psCodCta) Then
            MsgBox "El crédito tiene una Autorización pendiente", vbInformation, "Alerta"
            Call frmCredNewNivAutorizaVer.Consultar(psCodCta)
            Exit Sub
        End If
        Set oDCOMCredito = Nothing
        '***** Fin LUCV20171212 *****
        
        Set oCFRenNew = New COMDCartaFianza.DCOMCartaFianza
        Set rsCFRenNew = oCFRenNew.ObternerRenovacionAntCF(psCodCta, 2092)
        If Not (rsCFRenNew.BOF And rsCFRenNew.EOF) Then
            fnMotivo = CInt(rsCFRenNew!nMotivoRechazo)
            fdPrdEstado = CDate(rsCFRenNew!dPrdEstado)
        Else
            fnMotivo = 0
            fdPrdEstado = CDate(R!dAsignacion)
        End If
        
        fdVigCol = CDate(R!dVigCol)
        fdEmisionAnt = CDate(R!dAsignacion)
        fdVencAnt = CDate(R!dVencimiento)
        fnMontoRAnt = CDbl(R!nMontoCol)
        fnPrdEstado = CInt(R!nPrdEstado)
        
        Set rsCFRenNew = Nothing
        Set oCFRenNew = Nothing
        'WIOR FIN ************************************************
        
        lblCodigo.Caption = R!cPersCod
        lblNombre.Caption = PstaNombre(R!cPersNombre)
    
        lblCodAcreedor.Caption = R!cPersAcreedor
        lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
        lnModalidad = R!nModalidad
        'JOEP20181222 CP
        If R!nModalidad = 13 Then
            lblModalidad.Caption = R!OtrsModalidades
        Else
        'WIOR 20120611
        lblModalidad.Caption = R!sModalidad
        'MADM 20111020
        End If
        'JOEP20181222 CP
        lblCodAvalado.Caption = IIf(IsNull(R!cPersAvalado), "", R!cPersAvalado)
        If R!cAvalNombre <> "" Then
            lblNomAvalado.Caption = IIf(IsNull(PstaNombre(R!cAvalNombre)), "", PstaNombre(R!cAvalNombre))
        End If
        'END MADM
        'MAVM BAS II
        'If Mid(Trim(psCodCta), 9, 1) = "1" Then
            lblTipoCF = IIf(IsNull(R!cConsDescripcion), "", R!cConsDescripcion) '"COMERCIALES "
        'ElseIf Mid(Trim(psCodCta), 9, 1) = "2" Then
            'lblTipoCF = "MICROEMPRESA "
        'End If
        
        If Mid(Trim(psCodCta), 9, 1) = "1" Then
            lblmoneda = "Soles"
        ElseIf Mid(Trim(psCodCta), 9, 1) = "2" Then
            lblmoneda = "Dolares"
        End If
        lblanalista.Caption = PstaNombre(IIf(IsNull(R!cAnalista), "", R!cAnalista))
        'TxtMonApr.Text = IIf(IsNull(R!nMontoApr), "", Format(R!nMontoApr, "#0.00")) 'WIOR 20130313 COMENTO
        lblFecVencCF.Caption = IIf(IsNull(R!dVencimiento), "", Format(R!dVencimiento, "dd/mm/yyyy"))
        LblRenovacion.Caption = IIf(IsNull(R!nRenovacion), "0", R!nRenovacion)
        FraDatos.Enabled = True
        cmdGrabar.Enabled = True
        'By Capi Acta 035-2007
        'TxtFecEmiNue.Text = CDate(lblFecVencCF.Caption) + 1''WIOR 20130313 COMENTO
        lblPoliza.Caption = IIf(IsNull(R!nPoliza), "0", R!nPoliza)
        
        '**Fin************************************************************
        
        'End By
        'WIOR 20130313 *******************************************
        TxtFecEmiNue.Text = Format(ldNEmision, "dd/mm/yyyy")
        TxtMonApr.Text = Format(lnMontoR, "#0.00")
        TxtPeriodo.Text = lnPeriodo
        txtFecVencNueva.Text = Format(ldNVenc, "dd/mm/yyyy")
        
        TxtPeriodo.Locked = True
        TxtMonApr.Locked = True
        'WIOR FIN ************************************************
        cmdCheckListCFRenv.Enabled = True 'JOEP20190124 CP
    Else
        MsgBox "La Fecha de Vencimiento puede ser Menor o Mayor a la fecha que se desea Renovar", vbInformation, "AVISO"
        Exit Sub
    End If
Exit Sub

ErrorCargaDat:
    MsgBox "Error Nº [" & str(Err.Number) & "] " & Err.Description, vbCritical, "Error del Sistema"
    Exit Sub
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CargaDatosR (ActXCodCta.NroCuenta)
    End If
    'TxtPeriodo.SetFocus
End Sub

'JOEP20190124 CP
Private Sub cmdCheckListCFRenv_Click()
    If frmAdmCheckListDocument.Inicio(ActXCodCta.NroCuenta, 500, 514, CCur(Replace(TxtMonApr.Text, ",", "")), 0, nRegRenovacionCF) = True Then
        bCheckList = True
    Else
        bCheckList = False
    End If
End Sub
'JOEP20190124 CP

Private Sub cmdGrabar_Click()
Dim loNCartaFianza As COMNCartaFianza.NCOMCartaFianza
Dim loImprime As COMNCartaFianza.NCOMCartaFianzaImpre
Dim loPrevio As previo.clsprevio
Dim lsCadImprimir As String

Dim loContFunct As COMNContabilidad.NCOMContFunciones
Set loContFunct = New COMNContabilidad.NCOMContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnRechazo As Integer
Dim lsComenta As String
Dim lnRenovacion As Integer
Dim lnMonto As Double
Dim oDCredito As COMDCredito.DCOMCredito 'LUCV20171212, Agregó según observación SBS

'JOEP20190124 CP
    If bCheckList = False Then
        MsgBox "Debe registrar el CheckList", vbInformation, "Alerta"
        Exit Sub
    End If
'JOEP20190124 CP

lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Set loContFunct = Nothing

vCodCta = ActXCodCta.NroCuenta

If MsgBox("Desea Grabar la Renovación  de Carta Fianza", vbInformation + vbYesNo, "Aviso") = vbYes Then
    
    If txtFecVencNueva = "__/__/____" Then
        MsgBox "Falta calcular la nueva fecha de vencimiento.", vbInformation, "AVISO"
        Exit Sub
    End If
    
    lnRenovacion = CInt(LblRenovacion) + 1
    Set loNCartaFianza = New COMNCartaFianza.NCOMCartaFianza
        'Call loNCartaFianza.nCFRenovacion(vCodCta, txtFecVencNueva, lnRenovacion, gdFecSis)
        'By Capi Acta 035-2007 se modifico parametro TxtFecEmiNue y TxtMonApr
        
        'WIOR 20130313 *******************************************
        Set oCFRenNew = New COMDCartaFianza.DCOMCartaFianza
        Set rsCFRenNew = oCFRenNew.OperacionesCFRestaura(1, vCodCta, lnRenovacion - 1, CInt(TxtPeriodo.Text), fnMontoRAnt, fdVigCol, fdEmisionAnt, fdVencAnt, fnPrdEstado, fdPrdEstado, fnMotivo)
        Set oCFRenNew = Nothing
        Set rsCFRenNew = Nothing
        'WIOR FIN *************************************************
        Call loNCartaFianza.nCFRenovacion(vCodCta, txtFecVencNueva, lnRenovacion, TxtFecEmiNue, val(TxtMonApr))
        
        'MAVM 20101216 *** Historial de CF
        Call loNCartaFianza.nCFRegistraHistorial(vCodCta, lblFecVencCF.Caption, TxtPeriodo.Text, txtFecVencNueva.Text, TxtMonApr.Text)
        '***
        
        'MAVM 20100625 BAS II ***
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Renovacion de CF", vCodCta, gCodigoCuenta
        Set objPista = Nothing
        '***
    Set loNCartaFianza = Nothing
    
    MsgBox "Datos se guardaron satisfactoriamente.", vbInformation, "AVISO"
    
    'LUCV20171212, Según observacion SBS
    Set oDCredito = New COMDCredito.DCOMCredito
    If oDCredito.verificarExisteAutorizaciones(vCodCta) Then
        Call frmCredNewNivAutorizaVer.Consultar(vCodCta)
    End If
    Set oDCredito = Nothing
    'Fin LUCV20171212
    
    cmdGrabar.Enabled = False
    'cmdImprimir.Enabled = True
    FraDatos.Enabled = False
    
    LimpiarControles
    ActXCodCta.SetFocus
    
End If

'  Call RestMonto(vCodCta)

End Sub

Private Sub cmdSalir_Click()
    bCheckList = False 'JOEP20190124 CP
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    LimpiarControles
    gsOpeCod = gCredRenovacionCF
    
    cmdCheckListCFRenv.Enabled = False 'JOEP20190124 CP
End Sub

Sub LimpiarControles()
   ActXCodCta.Enabled = True
   ActXCodCta.NroCuenta = fgIniciaAxCuentaCF
   lblCodigo.Caption = ""
   lblNombre.Caption = ""
   lblCodAcreedor.Caption = ""
   lblNomAcreedor.Caption = ""
   lblCodAvalado.Caption = ""
   lblNomAvalado.Caption = ""
   lblTipoCF.Caption = ""
   lblmoneda.Caption = ""
   TxtMonApr.Text = ""
   lblModalidad.Caption = ""
   lblanalista.Caption = ""
   lblFecVencCF.Caption = ""
   LblRenovacion.Caption = ""
   
   '*** PEAC 20090903
   lblPoliza.Caption = ""
   lblcomision.Caption = ""
   txtFecVencNueva.Text = "__/__/____"
   TxtPeriodo.Text = " "
   TxtFecEmiNue.Text = "__/__/____"
   '*** FIN PEAC
   
   FraDatos.Enabled = False
   cmdGrabar.Enabled = False
   TxtMonApr.Enabled = True
    'WIOR 20130313 *******************************************
    fnMontoRAnt = 0
    fdVigCol = "01/01/1900"
    fdEmisionAnt = "01/01/1900"
    fdVencAnt = "01/01/1900"
    fnPrdEstado = 0
    fdPrdEstado = "01/01/1900"
    fnMotivo = 0
    'WIOR FIN *************************************************
End Sub

Private Sub txtFecVencNueva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub

Private Sub TxtPeriodo_Change()
If lblFecVencCF = "" Then Exit Sub 'JOEP20181227 CP
    If IsNumeric(IIf(TxtPeriodo.Text = "", 0, TxtPeriodo.Text)) Then
        txtFecVencNueva.Text = CDate(lblFecVencCF.Caption) + CInt(TxtPeriodo.Text) + 1
    End If
    Set loParam = New COMDColocPig.DCOMColPCalculos
    fpComision = loParam.dObtieneColocParametro(4001)
    Set loParam = Nothing
    sCodCta = (ActXCodCta.NroCuenta)

    Set lcCons = New COMDConstSistema.DCOMConstSistema
    Set lr = lcCons.ObtenerVarSistema()
        fbComisionTrimestral = IIf(lr!nConsSisValor = 2, True, False)
    Set lr = Nothing
    Set lcCons = Nothing
    
    If val(TxtMonApr.Text) > 0 Then '*** PEAC 20090930
        '**Inicio,Capi Octubre2007, mostrar comisión de carta fianza
        '**Las siguientes lineas de código fueron obtenidos de la pantalla de sugerenia
        Set loCFCalculo = New COMNCartaFianza.NCOMCartaFianzaCalculos
        If fbComisionTrimestral = False Then ' Caja Trujillo
            lblcomision = Format(loCFCalculo.nCalculaComisionCF(val(TxtMonApr.Text), DateDiff("d", CDate(TxtFecEmiNue), CDate(txtFecVencNueva)), fpComision, Mid(sCodCta, 9, 1)), "#,##0.00")
        Else  ' Caja Metropolitana
            lblcomision = Format(loCFCalculo.nCalculaComisionTrimestralCF(val(TxtMonApr.Text), DateDiff("d", CDate(TxtFecEmiNue), CDate(txtFecVencNueva)), lnModalidad, Mid(Trim(sCodCta), 9, 1), ActXCodCta.NroCuenta, 6), "#,###0.00")
        End If
        Set loCFCalculo = Nothing
    End If
End Sub

Private Sub TxtPeriodo_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii) 'JOEP20181222 CP
    If KeyAscii = 13 Then
'        If IsNumeric(TxtPeriodo) Then
'            txtFecVencNueva.Text = CDate(lblFecVencCF.Caption) + CInt(TxtPeriodo.Text)
'            cmdGrabar.SetFocus
'        Else
'            txtFecVencNueva.Text = "__/__/____"
'            MsgBox "Solo Ingrese Valores Numericos", vbInformation, "Aviso"
'            TxtPeriodo.Text = ""
'        End If
        
   If IsNumeric(TxtPeriodo) Then
        txtFecVencNueva.Text = CDate(lblFecVencCF.Caption) + CInt(TxtPeriodo.Text) + 1 'WIOR 20130313
        cmdGrabar.SetFocus
    ElseIf Len(TxtPeriodo) = 0 Then
        txtFecVencNueva.Text = "__/__/____"
        cmdGrabar.SetFocus
    ElseIf Not IsNumeric(TxtPeriodo) Then
        MsgBox "Solo Ingrese Valores Numéricos", vbInformation, "Aviso"
        TxtPeriodo.Text = ""
    End If
        
    End If
End Sub

Private Sub TxtPeriodo_LostFocus()

    If IsNumeric(TxtPeriodo) Then
        txtFecVencNueva.Text = CDate(lblFecVencCF.Caption) + CInt(TxtPeriodo.Text) + 1 'WIOR 20130313
        cmdGrabar.SetFocus
    ElseIf Len(TxtPeriodo) = 0 Then
        txtFecVencNueva.Text = "__/__/____"
        cmdGrabar.SetFocus
    ElseIf Not IsNumeric(TxtPeriodo) Then
        MsgBox "Solo Ingrese Valores Numericos", vbInformation, "Aviso"
        TxtPeriodo.Text = ""
    End If

End Sub
