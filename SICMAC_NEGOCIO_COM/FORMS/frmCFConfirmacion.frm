VERSION 5.00
Begin VB.Form frmCFConfirmacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta Fianza - Confirmación"
   ClientHeight    =   6210
   ClientLeft      =   2370
   ClientTop       =   405
   ClientWidth     =   7395
   Icon            =   "frmCFConfirmacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Avalado"
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
      Height          =   555
      Left            =   180
      TabIndex        =   24
      Top             =   1440
      Width           =   7050
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Avalado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblCodAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   26
         Tag             =   "txtcodigo"
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label lblNomAval 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2400
         TabIndex        =   25
         Tag             =   "txtnombre"
         Top             =   180
         Width           =   4470
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
      Height          =   555
      Left            =   180
      TabIndex        =   15
      Top             =   840
      Width           =   7050
      Begin VB.Label lblNomAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2400
         TabIndex        =   18
         Tag             =   "txtnombre"
         Top             =   180
         Width           =   4470
      End
      Begin VB.Label lblCodAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   17
         Tag             =   "txtcodigo"
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Acreedor "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame Frame6 
      Height          =   810
      Left            =   180
      TabIndex        =   12
      Top             =   5280
      Width           =   7035
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   5640
         TabIndex        =   29
         ToolTipText     =   "Ir al Menu Principal"
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   390
         Left            =   4320
         TabIndex        =   28
         ToolTipText     =   "Grabar Datos de Aprobacion de Credito"
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame fraDatos 
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
      Height          =   3075
      Left            =   180
      TabIndex        =   5
      Top             =   2160
      Width           =   7065
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
         Height          =   195
         Index           =   6
         Left            =   3840
         TabIndex        =   23
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento"
         Height          =   195
         Index           =   5
         Left            =   3840
         TabIndex        =   22
         Top             =   600
         Width           =   870
      End
      Begin VB.Label lblMontoApr 
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
         Height          =   315
         Left            =   4800
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblFecVencApr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4800
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Finalidad"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label lblAnalista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   4800
         TabIndex        =   14
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Analista"
         Height          =   195
         Index           =   1
         Left            =   3840
         TabIndex        =   13
         Top             =   1020
         Width           =   555
      End
      Begin VB.Label lblApoderado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   960
         TabIndex        =   0
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lblModalidad 
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
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblFinalidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   6735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Apoderado"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Modalidad"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   735
      End
      Begin VB.Label lblTipoCF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo "
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.Frame fraCliente 
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
      Height          =   600
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   7050
      Begin VB.Label lblNomcli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2400
         TabIndex        =   4
         Top             =   180
         Width           =   4485
      End
      Begin VB.Label lblCodcli 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   3
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   225
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmCFConfirmacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()
frmCFDarBaja.fbConfimacion = False
Unload Me
End Sub

Private Sub cmdConfirmar_Click()
If MsgBox("Esta seguro de confirmar la Baja de la Carta Fianza?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    frmCFDarBaja.fbConfimacion = True
    Hide
End If
End Sub

Private Sub Form_Load()
Call CentraForm(Me)
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Public Sub Inicio(ByVal psCtaCod As String)
Call CargaDatos(psCtaCod)
Me.Show 1
End Sub

Private Sub CargaDatos(ByVal psCta As String)

Dim oCF As COMDCartaFianza.DCOMCartaFianza 'DCartaFianza
Dim R As New ADODB.Recordset
Dim loCFCalculo As COMNCartaFianza.NCOMCartaFianzaCalculos 'NCartaFianzaCalculos
Dim loConstante As COMDConstantes.DCOMConstantes 'DConstante
Dim loCFValida As COMNCartaFianza.NCOMCartaFianzaValida 'NCartaFianzaValida
Dim lbTienePermiso As Boolean
Dim lnComisionPagada As Double
Dim lnComisionCalculada As Double
Dim ldFechaAsi As Date


    
    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set R = oCF.RecuperaCartaFianzaDarBaja(psCta)
    Set oCF = Nothing
    If Not R.BOF And Not R.EOF Then
        lblCodcli.Caption = R!cPersCod
        lblNomcli.Caption = PstaNombre(R!cPersNombre)
        lblCodAcreedor.Caption = R!cPersAcreedor
        lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
        ldFechaAsi = R!dAsignacion
        
        'MAVM 20100606
        'If Mid(Trim(psCta), 6, 1) = "1" Then
            lblTipoCF = IIf(IsNull(R!cConsDescripcion), "", R!cConsDescripcion) 'IIf(Mid(Trim(psCta), 9, 1) = "1", "COMERCIALES - SOLES", "COMERCIALES - DOLARES")
        'ElseIf Mid(Trim(psCta), 6, 1) = "2" Then
            'lblTipoCF = IIf(Mid(Trim(psCta), 9, 1) = "1", "MICROEMPRESA - SOLES", "MICROEMPRESA - DOLARES")
        'End If
        lblAnalista.Caption = IIf(IsNull(R!cAnalista), "", R!cAnalista)
        lblApoderado.Caption = IIf(IsNull(R!cApoderado), "", R!cApoderado)
        
        'MADM 20111020
        lblCodAvalado.Caption = IIf(IsNull(R!cAvalCod), "", R!cAvalCod)
        If (R!cAvalNombre) <> "" Then
            Me.lblNomAval.Caption = PstaNombre(R!cAvalNombre)
        End If
        'END MADM
        
        lblFinalidad.Caption = IIf(IsNull(R!cfinalidad), "", R!cfinalidad)
        lblMontoApr = IIf(IsNull(R!nMontoApr), "", Format(R!nMontoApr, "#0.00"))
        lblFecVencApr = IIf(IsNull(R!dVencApr), "", Format(R!dVencApr, "dd/mm/yyyy"))
        'fsEstado = R!nPrdEstado
        'fnRenovacion = IIf(IsNull(R!nRenovacion), 0, R!nRenovacion)
        'By Capi Acta 035-2007
        'If fsEstado <> gColocEstRenovada Then
        '    txtNumPoliza.Enabled = True
        'End If
        
        Set loConstante = New COMDConstantes.DCOMConstantes
            lblModalidad = loConstante.DameDescripcionConstante(gColCFModalidad, R!nModalidad)
        Set loConstante = Nothing
        'Verifica Fecha Vcto Aprobada es posterior a la Fecha Actual
        'If Format(lblFecVencApr, "yyyy/mm/dd") < Format(gdFecSis, "yyyy/mm/dd") Then
        '    MsgBox "Fecha de Vencimiento de Carta Fianza es anterior a la Fecha Actual", vbInformation, "Aviso"
            'LimpiaDatos
        '    Exit Sub
        'End If
        
        'Set loCFValida = New COMNCartaFianza.NCOMCartaFianzaValida
            'By capi 24022009 se modifico para enviar la fecha de proceso
            'lnComisionPagada = loCFValida.nCFPagoComision(psCta)
            'lnComisionPagada = loCFValida.nCFPagoComision(psCta, gdFecSis)
            '
        'Set loCFValida = Nothing
        'Set loCFCalculo = New COMNCartaFianza.NCOMCartaFianzaCalculos
            'lnComisionCalculada = Format(loCFCalculo.nCalculaComisionTrimestralCF(R!nMontoApr, DateDiff("d", ldFechaAsi, R!dVencApr), fpComision, Mid(psCta, 9, 1)), "####0.00")
        'Set loCFCalculo = Nothing

        '*** PEAC 20090813
        'If lnComisionPagada < lnComisionCalculada Then
        'If lnComisionPagada <= 0 Then
         '   MsgBox "No se ha pagado comision de Carta Fianza", vbInformation, "Aviso"
            'LimpiaDatos
          '  Exit Sub
        'End If
                
        'txtMontoApr.Text = IIf(IsNull(R!nMontoSug), "", Format(R!nMontoSug, "#0.00"))
        'TxtFecVenApr.Text = IIf(IsNull(R!dVencSug), "", Format(R!dVencSug, "dd/mm/yyyy"))
    
        fraDatos.Enabled = True
        'cmdGrabar.Enabled = True
        'cmdGenerarPDF.Enabled = True 'WIOR 20120613
    
    End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
frmCFDarBaja.fbConfimacion = False
End Sub
