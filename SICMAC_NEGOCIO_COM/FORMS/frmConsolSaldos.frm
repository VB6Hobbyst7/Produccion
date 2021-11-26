VERSION 5.00
Begin VB.Form frmConsolSaldos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saldos Consolidados"
   ClientHeight    =   4605
   ClientLeft      =   3270
   ClientTop       =   2760
   ClientWidth     =   6585
   Icon            =   "frmConsolSaldos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   330
      Left            =   5400
      TabIndex        =   15
      Top             =   4200
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saldos Consolidados"
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
      Height          =   2400
      Left            =   90
      TabIndex        =   8
      Top             =   1725
      Width           =   6450
      Begin VB.Label lblSaldoGarantizaTS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4950
         TabIndex        =   30
         Top             =   1520
         Width           =   1365
      End
      Begin VB.Label lblSaldoGarantizaD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3435
         TabIndex        =   29
         Top             =   1520
         Width           =   1365
      End
      Begin VB.Label lblSaldoGarantizaS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   28
         Top             =   1520
         Width           =   1365
      End
      Begin VB.Label Label9 
         Caption         =   "Créditos Garantiza:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   27
         Top             =   1520
         Width           =   1650
      End
      Begin VB.Label lblSaldoCFTS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4950
         TabIndex        =   24
         Top             =   1870
         Width           =   1365
      End
      Begin VB.Label lblSaldoColocTS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4950
         TabIndex        =   23
         Top             =   1155
         Width           =   1365
      End
      Begin VB.Label lblSaldoCaptacTS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4950
         TabIndex        =   22
         Top             =   795
         Width           =   1365
      End
      Begin VB.Label lblSaldoCFD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3435
         TabIndex        =   21
         Top             =   1870
         Width           =   1365
      End
      Begin VB.Label lblSaldoColocD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3435
         TabIndex        =   20
         Top             =   1155
         Width           =   1365
      End
      Begin VB.Label lblSaldoCaptacD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3435
         TabIndex        =   19
         Top             =   795
         Width           =   1365
      End
      Begin VB.Label lblSaldoCFS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Top             =   1870
         Width           =   1365
      End
      Begin VB.Label lblSaldoColocS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   1155
         Width           =   1365
      End
      Begin VB.Label lblSaldoCaptacS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Top             =   795
         Width           =   1365
      End
      Begin VB.Label Label7 
         Caption         =   "Total Soles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5055
         TabIndex        =   14
         Top             =   435
         Width           =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "Dolares Expres. en Soles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3465
         TabIndex        =   13
         Top             =   345
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "Soles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2235
         TabIndex        =   12
         Top             =   435
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Créditos Indirectos: (Cartas Fianzas)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1725
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Créditos Directos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   1155
         Width           =   1530
      End
      Begin VB.Label Label2 
         Caption         =   "Depósitos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   825
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
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
      Height          =   1590
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   6420
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   360
         Width           =   570
      End
      Begin VB.Label lblDocNatural 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Natural :"
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   1155
         Width           =   990
      End
      Begin VB.Label lblDocJuridico 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Jurídico :"
         Height          =   195
         Left            =   3585
         TabIndex        =   5
         Top             =   1170
         Width           =   1050
      End
      Begin VB.Label lblNomPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1230
         TabIndex        =   4
         Top             =   690
         Width           =   4890
      End
      Begin VB.Label lblDocNat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1230
         TabIndex        =   3
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label lblDocJur 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4725
         TabIndex        =   2
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label LblPersCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1230
         TabIndex        =   1
         Top             =   300
         Width           =   1755
      End
   End
   Begin VB.Label lblTipoCambio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1575
      TabIndex        =   26
      Top             =   4245
      Width           =   840
   End
   Begin VB.Label Label8 
      Caption         =   "Tipo de Cambio "
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
      Height          =   240
      Left            =   150
      TabIndex        =   25
      Top             =   4260
      Width           =   1425
   End
End
Attribute VB_Name = "frmConsolSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdsalir_Click()
    Unload Me
End Sub

Public Sub ConsolidaDatos(ByVal psCodPers As String, ByVal psNomPers As String, ByVal psDocId As String, _
                          ByVal psRUC As String)
        
'Dim oDatos As COMDCredito.DCOMCreditos
'Dim oTipoCambio As COMDConstSistema.NCOMTipoCambio
Dim lnTipoCambio As Currency

Dim lnSaldoCaptacS As Currency, lnSaldoCaptacD As Currency
Dim lnSaldoColocS As Currency, lnSaldoColocD As Currency
Dim lnSaldoCFS As Currency, lnSaldoCFD As Currency
Dim lnSaldoCaptacTS As Currency
Dim lnSaldoColocTS As Currency
Dim lnSaldoCFTS As Currency
Dim lnSaldoColocCFGarantizaS As Currency, lnSaldoColocCFGarantizaD As Currency, lnSaldoColocCFGarantizaTS As Currency 'EJVG20120430
'WIOR 20130122 *******************
Dim objPista As COMManejador.Pista
gsOpeCod = gsPersSaldosConsol
'WIOR FIN ************************
Dim oCred As COMDCredito.DCOMCredito

lblPersCod = psCodPers
lblNomPers = psNomPers
lblDocnat = psDocId
LblDocJur = psRUC

'Set oTipoCambio = New COMDConstSistema.NCOMTipoCambio
'    lnTipoCambio = oTipoCambio.EmiteTipoCambio(gdFecSis, TCFijoMes)
'    lblTipoCambio = Format(lnTipoCambio, "###,###,###.00  ")
'Set oTipoCambio = Nothing

'Set oDatos = New COMDCredito.DCOMCreditos
Set oCred = New COMDCredito.DCOMCredito
Call oCred.ConsolidaDatosCliente(gdFecSis, psCodPers, lnTipoCambio, lnSaldoCaptacS, lnSaldoCaptacD, lnSaldoColocS, lnSaldoColocD, _
                                lnSaldoCFS, lnSaldoCFD, lnSaldoCaptacTS, lnSaldoColocTS, lnSaldoCFTS, lnSaldoColocCFGarantizaS, lnSaldoColocCFGarantizaD, lnSaldoColocCFGarantizaTS)
                                'lnSaldoCFS, lnSaldoCFD, lnSaldoCaptacTS, lnSaldoColocTS, lnSaldoCFTS)
Set oCred = Nothing
'WIOR 20130122 ************************************
Set objPista = New COMManejador.Pista
objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gConsultar, "Consulta de Saldos Consolidados", psCodPers, gCodigoPersona
Set objPista = Nothing
'WIOR FIN *****************************************
    
lblTipoCambio = Format(lnTipoCambio, "###,###,##0.000  ")
'lnSaldoCaptacS = oDatos.SaldoConsolCapta(psCodPers, 1)
lblSaldoCaptacS = Format(lnSaldoCaptacS, "###,###,##0.00  ")
'lnSaldoCaptacD = oDatos.SaldoConsolCapta(psCodPers, 2)
lblSaldoCaptacD = Format(Round((lnSaldoCaptacD * lnTipoCambio), 2), "###,###,##0.00  ")

'lnSaldoColocS = oDatos.SaldoConsolColoc(psCodPers, 1)
lblSaldoColocS = Format(lnSaldoColocS, "###,###,##0.00  ")
'lnSaldoColocD = oDatos.SaldoConsolColoc(psCodPers, 2)
lblSaldoColocD = Format(Round((lnSaldoColocD * lnTipoCambio), 2), "###,###,##0.00  ")

'lnSaldoCFS = oDatos.SaldoConsolCF(psCodPers, 1)
lblSaldoCFS = Format(lnSaldoCFS, "###,###,##0.00  ")
'lnSaldoCFD = oDatos.SaldoConsolCF(psCodPers, 2)
lblSaldoCFD = Format(Round((lnSaldoCFD * lnTipoCambio), 2), "###,###,##0.00  ")

'lnSaldoCaptacTS = Round((lnSaldoCaptacS + (lnTipoCambio * lnSaldoCaptacD)), 2)
lblSaldoCaptacTS = Format(lnSaldoCaptacTS, "###,###,##0.00  ")
'lnSaldoColocTS = Round((lnSaldoColocS + (lnTipoCambio * lnSaldoColocD)), 2)
lblSaldoColocTS = Format(lnSaldoColocTS, "###,###,##0.00  ")
'lnSaldoCFTS = Round((lnSaldoCFS + (lnTipoCambio * lnSaldoCFD)), 2)
lblSaldoCFTS = Format(lnSaldoCFTS, "###,###,##0.00  ")
'EJVG20120430
lblSaldoGarantizaS.Caption = Format(lnSaldoColocCFGarantizaS, "###,###,##0.00  ")
lblSaldoGarantizaD.Caption = Format(lnSaldoColocCFGarantizaD * lnTipoCambio, "###,###,##0.00  ")
lblSaldoGarantizaTS.Caption = Format(lnSaldoColocCFGarantizaTS, "###,###,##0.00  ")
'Set oDatos = Nothing

Me.Show 1

End Sub
Private Sub Form_Load()
    CentraForm Me
End Sub
