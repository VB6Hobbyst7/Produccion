VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCFSugerencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta Fianza - Sugerencia de Analista"
   ClientHeight    =   6960
   ClientLeft      =   1545
   ClientTop       =   1635
   ClientWidth     =   8070
   Icon            =   "frmCFSugerencia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdVinculados 
      Caption         =   "Vinculados"
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
      Left            =   5280
      TabIndex        =   40
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Height          =   705
      Left            =   120
      TabIndex        =   22
      Top             =   6240
      Width           =   7815
      Begin VB.CommandButton cmdCheckListCF 
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
         Height          =   405
         Left            =   4200
         TabIndex        =   49
         ToolTipText     =   "CheckList"
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   405
         Left            =   6060
         TabIndex        =   5
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   405
         Left            =   1680
         TabIndex        =   4
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   405
         Left            =   360
         TabIndex        =   3
         Top             =   180
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   7815
      Begin VB.Frame Frame3 
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
         Height          =   615
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   7590
         Begin VB.Label lblNomAvalado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2280
            TabIndex        =   39
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   5130
         End
         Begin VB.Label lblCodAvalado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   38
            Tag             =   "txtcodigo"
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Avalado"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Sugerencia"
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
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   4680
         Width           =   7575
         Begin MSMask.MaskEdBox TxtFecVenSug 
            Height          =   300
            Left            =   6060
            TabIndex        =   2
            Top             =   225
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtMontoSug 
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
            ForeColor       =   &H80000012&
            Height          =   315
            Left            =   1725
            TabIndex        =   1
            Text            =   "0.00"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Monto Sugerido "
            Height          =   195
            Left            =   75
            TabIndex        =   21
            Top             =   285
            Width           =   1170
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Vencimiento Sugerido"
            Height          =   195
            Left            =   3900
            TabIndex        =   20
            Top             =   285
            Width           =   2040
         End
      End
      Begin VB.Frame fraLineaCred 
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
         Height          =   2535
         Left            =   120
         TabIndex        =   15
         Top             =   2100
         Width           =   7590
         Begin VB.Frame frAnalSol 
            Height          =   855
            Left            =   120
            TabIndex        =   44
            Top             =   1560
            Width           =   4590
            Begin VB.Label lblAnalista 
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
               Height          =   300
               Left            =   885
               TabIndex        =   48
               Top             =   120
               Width           =   3600
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Analista "
               ForeColor       =   &H80000001&
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   180
               Width           =   600
            End
            Begin VB.Label lblFecAsig 
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
               Height          =   300
               Left            =   900
               TabIndex        =   46
               Top             =   480
               Width           =   1290
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Solicitado"
               Height          =   195
               Left            =   120
               TabIndex        =   45
               Top             =   540
               Width           =   690
            End
         End
         Begin VB.Frame frModOtrsSeg 
            Height          =   615
            Left            =   120
            TabIndex        =   41
            Top             =   900
            Width           =   4575
            Begin VB.TextBox txtModOtrsSeg 
               Height          =   285
               Left            =   1080
               TabIndex        =   42
               Top             =   240
               Width           =   3375
            End
            Begin VB.Label lblModOtrs 
               Caption         =   "Modalidad Otros"
               Height          =   375
               Left            =   120
               TabIndex        =   43
               Top             =   165
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdActSubTipoCred 
            Height          =   315
            Left            =   4200
            Picture         =   "frmCFSugerencia.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Generar SubTipo Credito"
            Top             =   240
            Width           =   390
         End
         Begin VB.ComboBox cmbInstitucionFinanciera 
            Height          =   315
            Left            =   5160
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   240
            Width           =   2295
         End
         Begin VB.ComboBox cmbSubTpoCredCF 
            Height          =   315
            Left            =   1005
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label lblInstitucionFinanciera 
            AutoSize        =   -1  'True
            Caption         =   "IF: "
            Height          =   195
            Left            =   4920
            TabIndex        =   33
            Top             =   240
            Width           =   225
         End
         Begin VB.Label lblComision 
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
            Height          =   300
            Left            =   5940
            TabIndex        =   30
            Top             =   1680
            Width           =   1470
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Comision "
            Height          =   195
            Left            =   4920
            TabIndex        =   29
            Top             =   1680
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento"
            Height          =   195
            Left            =   4920
            TabIndex        =   28
            Top             =   1380
            Width           =   870
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Modalidad "
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   660
            Width           =   780
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Monto"
            Height          =   195
            Left            =   4920
            TabIndex        =   26
            Top             =   1020
            Width           =   450
         End
         Begin VB.Label lblMontoSol 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   5940
            TabIndex        =   25
            Top             =   960
            Width           =   1470
         End
         Begin VB.Label lblModCF 
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
            Height          =   300
            Left            =   1005
            TabIndex        =   24
            Top             =   600
            Width           =   3600
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
            Height          =   300
            Left            =   5940
            TabIndex        =   23
            Top             =   1320
            Width           =   1470
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   4920
            TabIndex        =   17
            Top             =   660
            Width           =   585
         End
         Begin VB.Label LblMoneda 
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
            Height          =   300
            Left            =   5940
            TabIndex        =   16
            Top             =   600
            Width           =   1470
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
         Height          =   615
         Left            =   135
         TabIndex        =   11
         Top             =   840
         Width           =   7590
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Acreedor "
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   12
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   5130
         End
      End
      Begin VB.Frame fracliente 
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
         Height          =   585
         Left            =   135
         TabIndex        =   7
         Top             =   195
         Width           =   7605
         Begin VB.Label lblNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2280
            TabIndex        =   10
            Tag             =   "txtnombre"
            Top             =   210
            Width           =   5160
         End
         Begin VB.Label lblCodigo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   9
            Tag             =   "txtcodigo"
            Top             =   210
            Width           =   1185
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Afianzado"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   705
         End
      End
   End
   Begin VB.CommandButton cmdExaminar 
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
      Left            =   6600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   390
      Left            =   180
      TabIndex        =   31
      Top             =   120
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   688
      Texto           =   "Cta Fianza"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmCFSugerencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*  APLICACION : Carta Fianza
'*  ARCHIVO : frmCFSolicitud
'*  CREACION: 01/09/2002    - LAYG
'*  MODIFICACION
'***************************************************************************
'*  RESUMEN: PERMITE REGISTRAR LOS DATOS SUGERIDOS POR EL ANALISTA
'***************************************************************************
Option Explicit

Dim vCodCta As String
Dim DatosCargados As Boolean
Dim fpComision As Double
Dim fbComisionTrimestral  As Boolean
Dim nActualizaSubTipoCred As Integer
Dim objPista As COMManejador.Pista
Dim fvGravamen() As tGarantiaGravamen 'EJVG20150715
Dim bCheckList As Boolean 'JOEP20190124 CP

Public Sub Inicia(Optional ByVal psCtaCF As String)
    bCheckList = False 'JOEP20190124 CP
    Call LimpiarSug
    Call CargaParametros
        
    If Len(psCtaCF) > 0 Then
        ActXCodCta.NroCuenta = psCtaCF
        Call CargaDatos(psCtaCF)
        DatosCargados = True
    Else
        DatosCargados = False
    End If
    Me.Show 1
End Sub

'PROCEDIMIENTO QUE CARGA LOS DATOS QUE SE REQUIEREN PARA EL FORMULARIO
Private Sub CargaDatos(ByVal psCta As String)
Dim oCF As COMDCartaFianza.DCOMCartaFianza 'DCartaFianza
Dim R As New ADODB.Recordset
Dim loCFCalculo As COMNCartaFianza.NCOMCartaFianzaCalculos 'NCartaFianzaCalculos
Dim loConstante As COMDConstantes.DCOMConstantes 'DConstante
    
    bCheckList = False 'JOEP20190124 CP
    
    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set R = oCF.RecuperaCartaFianzaSugerencia(psCta)
    Set oCF = Nothing
    If Not R.BOF And Not R.EOF Then
        lblCodigo.Caption = R!cPersCod
        lblNombre.Caption = PstaNombre(R!cPersNombre)

        lblCodAcreedor.Caption = R!cPersAcreedor
        lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
        
        lblCodAvalado.Caption = IIf(IsNull(R!cPersAvalado), "", R!cPersAvalado) 'MADM 20111020
                
        If R!cAvalNombre <> "" Then
            lblNomAvalado.Caption = PstaNombre(R!cAvalNombre)
        End If
        
'        If Mid(Trim(psCta), 6, 1) = "1" Then
'            lblTipoCartaF = "COMERCIALES"
'        ElseIf Mid(Trim(psCta), 6, 1) = "2" Then
'            lblTipoCartaF = "MICROEMPRESA"
'        End If
        
        If Mid(Trim(psCta), 9, 1) = "1" Then
            lblMoneda = "Soles"
        ElseIf Mid(Trim(psCta), 9, 1) = "2" Then
            lblMoneda = "Dolares"
        End If
        
        lblMontoSol = Format(Trim(R!nMontoSol), "#,###0.00")
        lblFecAsig = Format(Trim(R!dAsignacion), "dd/mm/yyyy")
        lblFecVencCF = Format(Trim(R!dVencSol), "dd/mm/yyyy")
        lblAnalista = IIf(IsNull(R!cAnalista), "", R!cAnalista)
        
        Call CargarSubTpoCred
        Call CargaInstitucionesFinancieras(gTpoInstFinanc)
        cmbSubTpoCredCF.ListIndex = IndiceListaCombo(cmbSubTpoCredCF, IIf(R!cTpoCredCod = "", 0, R!cTpoCredCod))
        
        If R!cTpoCredCod <> "" Then
            DatosCargados = True
        Else
            DatosCargados = False
        End If
                
        If IIf(IsNull(R!cTpoCredCod), 0, R!cTpoCredCod) = 181 Then
            cmbInstitucionFinanciera.Visible = True
            lblInstitucionFinanciera.Visible = True
            cmbInstitucionFinanciera.ListIndex = IndiceListaCombo(cmbInstitucionFinanciera, IIf(IsNull(R!nTpoInstCorp), 0, R!nTpoInstCorp))
        Else
            cmbInstitucionFinanciera.Visible = False
            lblInstitucionFinanciera.Visible = False
        End If
        
        TxtMontoSug.Text = IIf(IsNull(R!nMontoSug), "0.00", Format(Trim(R!nMontoSug), "#,###0.00"))
        TxtFecVenSug.Text = IIf(IsNull(R!dVencSug), "__/__/____", Format(Trim(R!dVencSug), "dd/mm/yyyy"))
        Set loConstante = New COMDConstantes.DCOMConstantes
            lblModCF = loConstante.DameDescripcionConstante(gColCFModalidad, R!nModalidad)
        'JOEP20181227 CP
            If R!nModalidad = 13 Then
                frmCFSugerencia.Height = 7380
                Frame5.top = 6240
                Frame1.Height = 5535
                Frame4.top = 4680
                fraLineaCred.Height = 2535
                frModOtrsSeg.Visible = True
                frAnalSol.BorderStyle = 0
                frAnalSol.top = 1560
                txtModOtrsSeg.Text = R!OtrsModalidades
            Else
                frModOtrsSeg.Visible = False
                frAnalSol.BorderStyle = 0
                frAnalSol.top = 960
                txtModOtrsSeg.Text = ""
            End If
        'JOEP20181227 CP
        Set loConstante = Nothing
            
        Set loCFCalculo = New COMNCartaFianza.NCOMCartaFianzaCalculos
            lblcomision = Format(loCFCalculo.nCalculaComisionCF(R!nMontoSol, DateDiff("d", gdFecSis, R!dVencSol), fpComision, Mid(psCta, 9, 1)), "####0.00")
        Set loCFCalculo = Nothing
        
        Set loCFCalculo = New COMNCartaFianza.NCOMCartaFianzaCalculos
            If fbComisionTrimestral = False Then ' Caja Trujillo
                lblcomision = Format(loCFCalculo.nCalculaComisionCF(R!nMontoSol, DateDiff("d", gdFecSis, R!dVencSol), fpComision, Mid(psCta, 9, 1)), "#,##0.00")
            Else  ' Caja Metropolitana
                lblcomision = Format(loCFCalculo.nCalculaComisionTrimestralCF(R!nMontoSol, DateDiff("d", gdFecSis, R!dVencSol), R!nModalidad, Mid(Trim(psCta), 9, 1), psCta, 6), "#,###0.00")
            End If
        Set loCFCalculo = Nothing
        
        cmdGrabar.Enabled = True
        TxtMontoSug.Enabled = True
        TxtMontoSug.Visible = True
        TxtMontoSug = lblMontoSol
        cmdCheckListCF.Enabled = True
        'TxtMontoSug.SetFocus
        
    End If
    R.Close
    Set R = Nothing
    
End Sub

Sub LimpiarSug()
    ActXCodCta.Enabled = True
    ActXCodCta.EnabledCta = True
    ActXCodCta.NroCuenta = fgIniciaAxCuentaCF
    lblCodigo.Caption = ""
    lblNombre.Caption = ""
    lblCodAcreedor.Caption = ""
    lblNomAcreedor.Caption = ""
    lblCodAvalado.Caption = ""
    lblNomAvalado.Caption = ""
    'lblTipoCartaF.Caption = ""
    lblMoneda.Caption = ""
    lblMontoSol.Caption = ""
    lblModCF.Caption = ""
    lblAnalista.Caption = ""
    lblFecAsig.Caption = ""
    lblFecVencCF.Caption = ""
    lblcomision.Caption = ""
    TxtMontoSug.Text = ""
    TxtFecVenSug.Text = "__/__/____"
    cmdGrabar.Enabled = False
    txtModOtrsSeg.Enabled = False 'JOEP20181227 CP
    cmdCheckListCF.Enabled = False 'JOEP20190124 CP
End Sub

Function ValidaDatosOk() As Boolean
Dim loCFValida As COMNCartaFianza.NCOMCartaFianzaValida 'NCartaFianzaValida
Dim lsSQL As String
Dim lnValorGarantGrav As Double
Dim loGen As COMDConstSistema.DCOMGeneral
Dim lnTipoCambioFijo As Double
Dim sCad As String
Dim lsmensaje As String

Dim objCFValida As COMDCartaFianza.DCOMCartaFianza 'JOEP ERS047
Dim RsTpProducto As ADODB.Recordset 'JOEP ERS047

    ValidaDatosOk = False
    
    sCad = ValidaFecha(Me.TxtFecVenSug.Text)
    If sCad <> "" Then
        MsgBox sCad, vbInformation, "Aviso"
        TxtFecVenSug.SetFocus
        Exit Function
    End If
    If Me.TxtMontoSug <= 0 Then
        MsgBox "Ingrese monto Correcto de Carta Fianza", vbInformation, "Aviso"
        TxtMontoSug.SetFocus
        Exit Function
    End If
    
    If CDate(Format(TxtFecVenSug.Text, "yyyy/mm/dd")) <= CDate(Format(gdFecSis, "yyyy/mm/dd")) Then
        MsgBox "Fecha de Vencimiento no puede ser anterior a la fecha actual", vbInformation, "Aviso"
        TxtFecVenSug.SetFocus
        Exit Function
    End If
    
    If DateDiff("d", lblFecAsig, TxtFecVenSug) <= -1 Then
        MsgBox "Fecha de vencimiento debe ser posterior a la de Solicitud de Carta Fianza", vbInformation, "Aviso"
        TxtFecVenSug.SetFocus
        Exit Function
    End If
    
    'MAVM 20100616 BAS II
    If cmbSubTpoCredCF.ListIndex = -1 Then
        MsgBox "Seleccione el Tipo del Credito", vbInformation, "Aviso"
        cmbSubTpoCredCF.SetFocus
        Exit Function
    End If
    
    Set loGen = New COMDConstSistema.DCOMGeneral
        'lnTipoCambioFijo = loGen.EmiteTipoCambio(gdFecSis, TCFijoMes)
        lnTipoCambioFijo = loGen.EmiteTipoCambio(gdFecSis, TCFijoDia) 'EJVG20150713
    Set loGen = Nothing
    Set loCFValida = New COMNCartaFianza.NCOMCartaFianzaValida
        lnValorGarantGrav = loCFValida.nCFGarantiasGravada(vCodCta, lnTipoCambioFijo, lsmensaje)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Function
        End If
    Set loCFValida = Nothing
    
    If lnValorGarantGrav = 0 Then
        MsgBox "El crédito no cuenta con Garantías relacionadas", vbInformation, "Aviso"
        Exit Function
    ElseIf CDbl(Format(TxtMontoSug.Text, "#0.00")) > lnValorGarantGrav Then
        'VERIFICA QUE MONTO DE GARANTIAS SEA MAYOR QUE MONTO SUGERIDO
        If MsgBox("Monto de Garantias : " & Format(lnValorGarantGrav, "#0.00") & " es Menor " & _
                  "Que Monto Sugerido. Desea Continuar ", vbInformation + vbYesNo, "Aviso") = vbNo Then
            Exit Function
        End If
    End If
    
    'JOEP20190124 CP
    If bCheckList = False Then
        MsgBox "Debe registrar el CheckList", vbInformation, "Alerta"
        Exit Function
    End If
    'JOEP20190124 CP
    
    'JOEP ERS047 20170904
        Set objCFValida = New COMDCartaFianza.DCOMCartaFianza
        Set RsTpProducto = objCFValida.get_VerificaCFAutoLiqxHipot(ActXCodCta.NroCuenta, 17)
        Set objCFValida = Nothing
        If Not (RsTpProducto.EOF And RsTpProducto.BOF) Then
            If RsTpProducto!nEstado = 0 Then
                MsgBox "El crédito supera el porcentaje máximo de Carta Fianza con Garantia Autoliquidable. Cualquier consulta comunicarse con Riesgos.", vbInformation, "Aviso" 'WIOR 20150714
                Exit Function
            ElseIf RsTpProducto!nEstado = 2 Then
                MsgBox "El crédito no puede ser sugerido, pues su solicitud de autorización fue rechazada por la gerencia de riesgos", vbInformation, "Aviso"
                Exit Function
            End If
        End If
        RsTpProducto.Close
        
        Set objCFValida = New COMDCartaFianza.DCOMCartaFianza
        Set RsTpProducto = objCFValida.get_VerificaCFAutoLiqxHipot(ActXCodCta.NroCuenta, 8)
        Set objCFValida = Nothing
        If Not (RsTpProducto.EOF And RsTpProducto.BOF) Then
            If RsTpProducto!nEstado = 0 Then
                MsgBox "El crédito supera el porcentaje máximo de Carta Fianza con Garantia Hipotecaria. Cualquier consulta comunicarse con Riesgos.", vbInformation, "Aviso" 'WIOR 20150714
                Exit Function
            ElseIf RsTpProducto!nEstado = 2 Then
                MsgBox "El crédito no puede ser sugerido, pues su solicitud de autorización fue rechazada por la gerencia de riesgos", vbInformation, "Aviso"
                Exit Function
            End If
        End If
        RsTpProducto.Close
    'JOEP ERS047 20170904
    
    ValidaDatosOk = True
End Function

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    vCodCta = ActXCodCta.NroCuenta
    If Len(vCodCta) > 0 Then
        Call CargaDatos(vCodCta)
        ActXCodCta.Enabled = False
    Else
        Call LimpiarSug
        ActXCodCta.SetFocus
    End If
End Sub

Private Sub cmbSubTpoCredCF_Click()
    If Right(cmbSubTpoCredCF.Text, 3) = gColCredCorpoCF Then
        lblInstitucionFinanciera.Visible = True
        cmbInstitucionFinanciera.Visible = True
    Else
        lblInstitucionFinanciera.Visible = False
        cmbInstitucionFinanciera.Visible = False
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiarSug
    cmdGrabar.Enabled = False
    ActXCodCta.SetFocusProd
End Sub

'JOEP20190124 CP
Private Sub cmdCheckListCF_Click()
Dim objGar As COMDCartaFianza.DCOMCartaFianza
Dim rsGar As ADODB.Recordset

If cmbSubTpoCredCF.Text = "" Then
    MsgBox "Seleccione el Tipo de Crédito", vbInformation, "Aviso"
    If cmbSubTpoCredCF.Enabled = True Then
        cmbSubTpoCredCF.SetFocus
    End If
    Exit Sub
End If
Set objGar = New COMDCartaFianza.DCOMCartaFianza
Set rsGar = objGar.get_ValidadCobGar(ActXCodCta.NroCuenta)
If Not (rsGar.BOF And rsGar.EOF) Then
    If rsGar!cMsg <> "" Then
        MsgBox rsGar!cMsg, vbInformation, "Aviso"
        Exit Sub
    End If
End If
RSClose rsGar

Set objGar = New COMDCartaFianza.DCOMCartaFianza
Set rsGar = objGar.get_ValidadCF(ActXCodCta.NroCuenta, Right(cmbSubTpoCredCF.Text, 9), gsCodCargo)
If Not (rsGar.BOF And rsGar.EOF) Then
    If rsGar!cMensaje <> "" Then
        MsgBox rsGar!cMensaje, vbInformation, "Aviso"
        Exit Sub
    End If
End If

    If frmAdmCheckListDocument.Inicio(ActXCodCta.NroCuenta, 500, 514, CCur(Replace(TxtMontoSug.Text, ",", "")), Trim(Right(cmbSubTpoCredCF.Text, 9)), nRegSugerenciaCF) = True Then  'JOEP20181229 CP
        bCheckList = True
    Else
        bCheckList = False
    End If

Set objGar = Nothing
RSClose rsGar
End Sub
'JOEP20190124 CP

Private Sub cmdExaminar_Click()
Dim lsCta As String
    'MAVM 20100604 Se agrego la var: gColCFTpoProducto BAS II***
    lsCta = frmCFPersEstado.Inicio(Array(gColocEstSolic, gColocEstSug), "Sugerencia de Carta Fianza", Array(gColCFComercial, gColCFPYME, gColCFTpoProducto))
    If Len(Trim(lsCta)) > 0 Then
        ActXCodCta.NroCuenta = lsCta
        ActXCodCta.Enabled = False
        Call CargaDatos(lsCta)
    Else
        Call LimpiarSug
    End If
End Sub

Private Sub cmdGrabar_Click()
Dim loNCartaFianza As COMNCartaFianza.NCOMCartaFianza 'NCartaFianza
Dim loImprime As COMNCartaFianza.NCOMCartaFianzaReporte
Dim loPrevio As previo.clsprevio
Dim lsCadImprimir As String

Dim loContFunct As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnMontoSug As Currency
Dim ldVencSug As Date
Dim ldAsigSug As Date 'FRHU20131126

Dim lsmensaje As String

vCodCta = ActXCodCta.NroCuenta

If ValidaDatosOk = False Then
    Exit Sub
End If

lnMontoSug = Format(TxtMontoSug.Text, "#0.00")
ldVencSug = Format(TxtFecVenSug.Text, "dd/mm/yyyy")
ldAsigSug = Format(lblFecAsig, "dd/mm/yyyy") 'FRHU20131126

If Not RecalcularCoberturaGarantias(vCodCta, False, "514", "CARTA FIAMZA", CCur(TxtMontoSug.Text), fvGravamen) Then Exit Sub   'EJVG20150715
If MsgBox("Desea Guardar Sugerencia de Analista", vbInformation + vbYesNo, "Sugerencia de Analista") = vbYes Then
    'EJVG20150715 ***
    If Not IsArray(fvGravamen) Then
        ReDim fvGravamen(0)
    End If
    'END EJVG *******
    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
    Set loNCartaFianza = New COMNCartaFianza.NCOMCartaFianza
        'MAVM 20100605 Se agrego CInt(Trim(Right(cmbSubTpoCredCF.Text, 3)))
        'FRHU20131126 Se agrego el parametro: ldAsigSug
        Call loNCartaFianza.nCFSugerencia(vCodCta, lsFechaHoraGrab, ldAsigSug, ldVencSug, lnMontoSug, CInt(Trim(Right(cmbSubTpoCredCF.Text, 3))), Trim(Right(cmbInstitucionFinanciera, 3)), fvGravamen, Trim(txtModOtrsSeg.Text))
    Set loNCartaFianza = Nothing
    
    'MAVM 20100621
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, IIf(DatosCargados = False, gInsertar, gModificar), "Sugerencia CF", vCodCta, gCodigoCuenta
    
    ' *** Impresion
    If MsgBox(" Desea Imprimir Resumen de Comite para Carta Fianza ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        Set loImprime = New COMNCartaFianza.NCOMCartaFianzaReporte
            loImprime.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lsCadImprimir = loImprime.nRepoDuplicado(vCodCta, 3, lsmensaje, gImpresora)
            If Trim(lsmensaje) <> "" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Sub
            End If
        Set loImprime = Nothing
        Set loPrevio = New previo.clsprevio
            loPrevio.Show lsCadImprimir, "Carta Fianza- Sugerencia de Analista", True, , gImpresora
            'Do While True
            '    If MsgBox("Desea Imprimir Resumen de Comite para Carta Fianza ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            '        loPrevio.PrintSpool sLpt, lsCadImprimir, True
            '    Else
                    Set loPrevio = Nothing
            '        Exit Do
            '    End If
            'Loop
    End If
    frmCFHojaAprob.Inicio vCodCta
    Call LimpiarSug
    cmdGrabar.Enabled = False
    
    'If lbSegCred = True Then
        'Unload Me
        'Unload FrmGravarGarantia
        'Unload frmSolCred
        'Unload frmRefinanCred
    'End If
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
'ALPA20130715************************
Private Sub cmdVinculados_Click()
    frmGruposEconomicos.Show 1
End Sub
'************************************

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
lblInstitucionFinanciera.Visible = False
cmbInstitucionFinanciera.Visible = False
nActualizaSubTipoCred = 0
gsOpeCod = gCredSugerenciaCF 'MAVM 20100621

'joep20181221 CP
    frmCFSugerencia.Height = 6915
    Frame5.top = 5760
    Frame4.top = 4200
    Frame1.Height = 5055
    fraLineaCred.Height = 2100
    frModOtrsSeg.Visible = False
    frAnalSol.BorderStyle = 0
    frAnalSol.top = 960
    txtModOtrsSeg.Text = ""
'joep20181221 CP
End Sub

Private Sub TxtFecVenSug_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(TxtFecVenSug) Then

        If CDate(Format(TxtFecVenSug.Text, "yyyy/mm/dd")) < CDate(Format(gdFecSis, "yyyy/mm/dd")) Then
            MsgBox "Fecha de Vencimiento no puede ser anterior a la fecha actual", vbInformation, "Aviso"
            TxtFecVenSug.SetFocus
            Exit Sub
        Else
            cmdGrabar.Enabled = True
            cmdGrabar.SetFocus
        End If
    Else
        TxtFecVenSug.SetFocus
        Exit Sub
    End If
End If
End Sub

'Private Sub TxtFecVenSug_LostFocus()
'Dim sCad As String
'    sCad = ValidaFecha(TxtFecVenSug.Text)
'    If sCad <> "" Then
'        MsgBox sCad, vbInformation, "Aviso"
'        Exit Sub
'    End If
'    If CDate(Format(TxtFecVenSug.Text, "dd/mm/yyyy")) < CDate(Format(gdFecSis, "dd/mm/yyyy")) Then
'        MsgBox "Fecha de Vencimiento no puede ser anterior a la fecha actual", vbInformation, "Aviso"
'        TxtFecVenSug.SetFocus
'        Exit Sub
'    Else
'        CmdGrabar.Enabled = True
'        CmdGrabar.SetFocus
'    End If
'
'End Sub

Private Sub txtMontoSug_GotFocus()
    TxtMontoSug.SelStart = 0
    TxtMontoSug.SelLength = Len(TxtMontoSug.Text)
End Sub

Private Sub txtMontoSug_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(TxtMontoSug, KeyAscii, 10, 3)
If KeyAscii = 13 Then
    TxtFecVenSug.SetFocus
End If
End Sub

Private Sub txtMontoSug_LostFocus()
    If Len(Trim(TxtMontoSug.Text)) = 0 Then
        TxtMontoSug.Text = "0"
    End If
    TxtMontoSug.Text = Format(TxtMontoSug.Text, "#,#0.00")
End Sub

'Carga los Parametros
Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos 'DColPCalculos
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

'MAVM BAS II 20100607***
Private Sub cmdActSubTipoCred_Click()
Dim oDCredito As COMDCredito.DCOMCredito
Dim lnSubTipoCredito As Integer
    If nActualizaSubTipoCred = 0 Then
        nActualizaSubTipoCred = 1
        Set oDCredito = New COMDCredito.DCOMCredito
        lnSubTipoCredito = oDCredito.ObtenerTipoCreditoxTipificacion(Trim(lblCodigo.Caption))
        If lnSubTipoCredito <> "0" Then
            cmbSubTpoCredCF.ListIndex = IndiceListaCombo(cmbSubTpoCredCF, Mid(lnSubTipoCredito, 1, 1) & "81")
        End If
        Set oDCredito = Nothing
    Else
        If MsgBox("El proceso para determinar el subtipo de credito ya fue realizado, Desea volver a realizarlo ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            nActualizaSubTipoCred = 0
            Call cmdActSubTipoCred_Click
        End If
    End If
End Sub

Private Sub CargarSubTpoCred()
    Dim rs As ADODB.Recordset
    Dim oCred As COMDCredito.DCOMCredito
    Set oCred = New COMDCredito.DCOMCredito
        Set rs = oCred.RecuperaSubTpoCrediticioCF
   
    cmbSubTpoCredCF.Clear
    Do While Not rs.EOF
        cmbSubTpoCredCF.AddItem rs!cConsDescripcion & Space(250) & rs!nConsValor
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set oCred = Nothing
End Sub

Private Sub CargaInstitucionesFinancieras(ByVal psTipo As String)
Dim oCons As COMDConstantes.DCOMConstantes
Dim sSql As String
Dim RTemp As ADODB.Recordset
    On Error GoTo ERRORCargaInstitucionesFinancieras
    Set oCons = New COMDConstantes.DCOMConstantes
    Set RTemp = oCons.RecuperaConstantes(psTipo)
    Set oCons = Nothing
    cmbInstitucionFinanciera.Clear
    Do While Not RTemp.EOF
        cmbInstitucionFinanciera.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Call CambiaTamañoCombo(cmbInstitucionFinanciera, 250)
    Exit Sub
    
ERRORCargaInstitucionesFinancieras:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub
'***
