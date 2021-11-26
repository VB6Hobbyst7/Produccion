VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColPAdjudicaLotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio : Adjudicación de Lotes"
   ClientHeight    =   5235
   ClientLeft      =   1680
   ClientTop       =   2145
   ClientWidth     =   7035
   Icon            =   "frmColPAdjudicaLotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraContenedor 
      Caption         =   "Adjudicación de Lotes"
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
      Height          =   5175
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5580
         TabIndex        =   22
         Top             =   4635
         Width           =   975
      End
      Begin VB.Frame fraImpresion 
         Caption         =   "Impresión"
         Height          =   555
         Left            =   120
         TabIndex        =   23
         Top             =   4560
         Width           =   2370
         Begin VB.OptionButton optImpresion 
            Caption         =   "Excel"
            Height          =   270
            Index           =   2
            Left            =   1185
            TabIndex        =   46
            Top             =   195
            Width           =   990
         End
         Begin VB.OptionButton optImpresion 
            Caption         =   "Impresora"
            Height          =   270
            Index           =   1
            Left            =   3120
            TabIndex        =   21
            Top             =   285
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.OptionButton optImpresion 
            Caption         =   "Pantalla"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   20
            Top             =   225
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   1815
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   5190
         TabIndex        =   3
         Top             =   300
         Width           =   1410
      End
      Begin VB.TextBox txtNumAdjudica 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   13
         Top             =   1815
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4320
         TabIndex        =   12
         Top             =   1815
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame fraContenedor 
         Height          =   2295
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   2220
         Width           =   6465
         Begin VB.CommandButton cmdDetalladoAdj 
            Caption         =   "Listado Detallado de Adjudicaciones con SIAF"
            Height          =   360
            Left            =   225
            TabIndex        =   35
            Top             =   2655
            Visible         =   0   'False
            Width           =   4485
         End
         Begin VB.CommandButton cmdPlanPreviaAnt 
            Caption         =   "Planilla Previa de Adjudicaciones con SIAF"
            Height          =   360
            Left            =   240
            TabIndex        =   45
            Top             =   2280
            Visible         =   0   'False
            Width           =   4485
         End
         Begin VB.CommandButton cmPlanAnt 
            Caption         =   "Listado Consolidado de Adjudicaciones con SIAF"
            Height          =   360
            Left            =   240
            TabIndex        =   44
            Top             =   2520
            Visible         =   0   'False
            Width           =   4485
         End
         Begin VB.CommandButton cmdAgencia 
            Caption         =   "A&gencias..."
            Height          =   345
            Left            =   5040
            TabIndex        =   43
            Top             =   1815
            Width           =   1020
         End
         Begin VB.Frame Frame1 
            Caption         =   "Adjudicaciones"
            Height          =   1560
            Left            =   4800
            TabIndex        =   39
            Top             =   210
            Width           =   1485
            Begin VB.ListBox List1 
               Height          =   1230
               Left            =   105
               TabIndex        =   18
               Top             =   195
               Width           =   1275
            End
         End
         Begin VB.CommandButton cmdImpListAdju 
            Caption         =   "Listado Consolidado de Joyas Adjudicadas (Para Subasta)"
            Height          =   360
            Left            =   210
            TabIndex        =   16
            Top             =   1200
            Width           =   4485
         End
         Begin VB.CommandButton cmdImpPlanPrevAdju 
            Caption         =   "Planilla Previa de Adjudicaciones"
            Height          =   360
            Left            =   210
            TabIndex        =   14
            Top             =   330
            Width           =   4485
         End
         Begin VB.CommandButton cmdAdjuLote 
            Caption         =   "Adjudicar Lotes"
            Enabled         =   0   'False
            Height          =   360
            Left            =   210
            TabIndex        =   15
            Top             =   760
            Width           =   4485
         End
         Begin MSMask.MaskEdBox txtFecMes 
            Height          =   315
            Left            =   3450
            TabIndex        =   19
            Top             =   1440
            Visible         =   0   'False
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
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
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdImpListAdjuMes 
            Caption         =   "Listado de Joyas Adjudicadas (por mes)"
            Height          =   360
            Left            =   240
            TabIndex        =   17
            Top             =   1320
            Visible         =   0   'False
            Width           =   3165
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fecha a listar"
            Height          =   195
            Index           =   11
            Left            =   3465
            TabIndex        =   42
            Top             =   1275
            Visible         =   0   'False
            Width           =   990
         End
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Precio del Oro "
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
         Height          =   975
         Index           =   2
         Left            =   180
         TabIndex        =   25
         Top             =   705
         Width           =   6465
         Begin VB.TextBox txtTipoCambio 
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
            Height          =   285
            Left            =   5445
            MaxLength       =   7
            TabIndex        =   5
            Top             =   210
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox txtPreOroInternacional 
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
            Height          =   285
            Left            =   3810
            MaxLength       =   6
            TabIndex        =   4
            Top             =   210
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtPreOro21 
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
            Height          =   285
            Left            =   5520
            MaxLength       =   6
            TabIndex        =   9
            Top             =   570
            Width           =   750
         End
         Begin VB.TextBox txtPreOro18 
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
            Height          =   285
            Left            =   3885
            MaxLength       =   6
            TabIndex        =   8
            Top             =   570
            Width           =   750
         End
         Begin VB.TextBox txtPreOro16 
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
            Height          =   285
            Left            =   2340
            MaxLength       =   6
            TabIndex        =   7
            Top             =   600
            Width           =   750
         End
         Begin VB.TextBox txtPreOro14 
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
            Height          =   285
            Left            =   705
            MaxLength       =   6
            TabIndex        =   6
            Top             =   585
            Width           =   750
         End
         Begin MSMask.MaskEdBox txtFecCorte 
            Height          =   315
            Left            =   1920
            TabIndex        =   48
            Top             =   200
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Corte:"
            Height          =   195
            Index           =   12
            Left            =   720
            TabIndex        =   47
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cambio :"
            Height          =   195
            Index           =   5
            Left            =   4800
            TabIndex        =   38
            Top             =   255
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Precio Oro Internacional :"
            Height          =   195
            Index           =   0
            Left            =   3120
            TabIndex        =   37
            Top             =   240
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "14 K :"
            Height          =   225
            Index           =   1
            Left            =   225
            TabIndex        =   29
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "16 K :"
            Height          =   225
            Index           =   2
            Left            =   1845
            TabIndex        =   28
            Top             =   615
            Width           =   450
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "18 K :"
            Height          =   225
            Index           =   3
            Left            =   3390
            TabIndex        =   27
            Top             =   600
            Width           =   450
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "21 K :"
            Height          =   225
            Index           =   4
            Left            =   5040
            TabIndex        =   26
            Top             =   600
            Width           =   450
         End
      End
      Begin MSMask.MaskEdBox txtFecAdjudica 
         Height          =   315
         Left            =   2100
         TabIndex        =   1
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHorAdjudica 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Left            =   3795
         TabIndex        =   2
         Top             =   300
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ProgressBar prgList 
         Height          =   330
         Left            =   2520
         TabIndex        =   40
         Top             =   4680
         Visible         =   0   'False
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSMask.MaskEdBox txtFecIni 
         Height          =   315
         Left            =   1350
         TabIndex        =   10
         Top             =   1860
         Visible         =   0   'False
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha hasta Adjudicar :"
         Height          =   390
         Index           =   10
         Left            =   210
         TabIndex        =   41
         Top             =   1785
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Estado :"
         Height          =   255
         Index           =   7
         Left            =   4590
         TabIndex        =   34
         Top             =   345
         Width           =   645
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Número :"
         Height          =   255
         Index           =   9
         Left            =   165
         TabIndex        =   33
         Top             =   330
         Width           =   660
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Hora :"
         Height          =   255
         Index           =   6
         Left            =   3330
         TabIndex        =   32
         Top             =   330
         Width           =   525
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha :"
         Height          =   255
         Index           =   8
         Left            =   1530
         TabIndex        =   31
         Top             =   330
         Width           =   555
      End
   End
   Begin VB.OLE OleExcel 
      Class           =   "Excel.Sheet.8"
      Height          =   870
      Left            =   135
      OleObjectBlob   =   "frmColPAdjudicaLotes.frx":030A
      TabIndex        =   36
      Top             =   90
      Visible         =   0   'False
      Width           =   1800
   End
End
Attribute VB_Name = "frmColPAdjudicaLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* ADJUDICACION DE CONTRATOS PIGNORATICIOS
'Archivo:  frmColPAdjudicaLotes.frm
'LAYG   :  01/06/2001.
'Resumen:  Nos permite registrar el precio de oro con el que se va ha procesar
'          el listado de contratos Adjudicados
Option Explicit

Dim pPrevioMax As Double
Dim pLineasMax As Double
Dim pHojaFiMax As Integer

Dim fnVerCodAnt As Boolean

Dim fnTasaImpuesto As Double
Dim fnTasaCustodiaVencida As Double
Dim fnTasaPreparacionRemate As Double
Dim fnNroRematesAdjudicacion As Integer
Dim fsProcesoCadaAgencia As String

Dim vRTFImp As String
Dim vCont As Double
Dim vNomAge As String
'************************************************************************************
Dim vPage As Integer
Dim vTot14 As Currency, vTot16 As Currency, vTot18 As Currency, vTot21 As Currency
Dim vTotDeuda As Currency, vTotValReg As Currency, vTotPieza As Long
Dim vTotAge As Integer

Private Sub HabilitaControles(ByVal pbEditar As Boolean, ByVal pbGrabar As Boolean, ByVal pbSalir As Boolean, _
    ByVal pbCancelar As Boolean, ByVal pbFecAdjudica As Boolean, ByVal pbFecIni As Boolean, _
    ByVal pbHorAdjudica As Boolean, ByVal pbOroInternacional As Boolean, ByVal pbTipoCambio As Boolean, _
    ByVal pbImpPlanPrevAdju As Boolean, ByVal pbAdjuLote As Boolean, ByVal pbImpActaAdju As Boolean, _
    ByVal pbImpListAdju As Boolean, ByVal pbList1 As Boolean)

    cmdEditar.Enabled = pbEditar
    cmdGrabar.Enabled = pbGrabar
    cmdSalir.Enabled = pbSalir
    cmdCancelar.Enabled = pbCancelar

    txtFecAdjudica.Enabled = False 'pbFecAdjudica '*** PEAC 20080714
    txtFecIni.Enabled = pbFecIni
    txtHorAdjudica.Enabled = False 'pbHorAdjudica '*** PEAC 20080714

    txtFecCorte.Enabled = False

    'txtPreOroInternacional.Enabled = pbOroInternacional
    'txtTipoCambio.Enabled = pbTipoCambio
    txtPreOro14.Enabled = False 'pbOroInternacional '*** PEAC 20080714
    txtPreOro16.Enabled = False 'pbOroInternacional '*** PEAC 20080714
    txtPreOro18.Enabled = False 'pbOroInternacional '*** PEAC 20080714
    txtPreOro21.Enabled = False 'pbOroInternacional '*** PEAC 20080714
    cmdImpPlanPrevAdju.Enabled = pbImpPlanPrevAdju
    cmdAdjuLote.Enabled = pbAdjuLote
    cmdImpListAdju.Enabled = pbImpListAdju
    List1.Enabled = pbList1
End Sub

Private Sub cmdAdjuLote_Click()

On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loAdjud As COMNColoCPig.NCOMColPRecGar
Dim lsFecAdjud As String, lsFecParaAdjud As String
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim x As Integer

Dim lsmensaje As String
    
    lsFecAdjud = Format$(Me.txtFecAdjudica, "mm/dd/yyyy")
    lsFecParaAdjud = Format$(Me.txtFecIni, "mm/dd/yyyy")
    
    If MsgBox("Esta seguro de Realizar la Adjudicación de Creditos Pignoraticios ? " _
            & vbCr & " ", vbYesNo + vbQuestion + vbDefaultButton2, " Aviso ") = vbYes Then

        ' El proceso de Adjudicacion se realiza por Agencia
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
            
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        
        Set loAdjud = New COMNColoCPig.NCOMColPRecGar
            Call loAdjud.nAdjudicaLotes(lsMovNro, lsFechaHoraGrab, "A", Trim(txtNumAdjudica.Text), _
                0, 0, lsFecAdjud, _
                val(Me.txtPreOro14.Text), val(Me.txtPreOro16.Text), val(Me.txtPreOro18.Text), val(txtPreOro21.Text), _
                lsFecParaAdjud, fsProcesoCadaAgencia, fnNroRematesAdjudicacion, lsmensaje, Me.txtFecCorte.Text, gdFecSis)
                ''*** PEAC 20090401 SE ADICIONÓ EL PARAMETRO gdFecSis
                
                
                If Trim(lsmensaje) <> "" Then
                     MsgBox lsmensaje, vbInformation, "Aviso"
                     Exit Sub
                End If
        Set loAdjud = Nothing
        
        txtEstado = "TERMINADO"
        cmdEditar.Enabled = False
        cmdGrabar.Enabled = False
        cmdCancelar.Enabled = False
        cmdImpPlanPrevAdju.Enabled = False
        cmdPlanPreviaAnt.Enabled = False
        cmdAdjuLote.Enabled = False
        'cmdImpActaAdju.Enabled = True
        cmdImpListAdju.Enabled = True
        cmPlanAnt.Enabled = True
        
        CargaAdjudicaciones
        
    End If
    
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdAgencia_Click()
    frmSelectAgencias.inicio Me
    frmSelectAgencias.Show 1
End Sub

'Permite no reconocer los datos ingresados
Private Sub cmdCancelar_Click()
    Call HabilitaControles(True, False, True, False, False, False, False, False, False, False, False, False, False, True)
    CargaDatosAdjudicacion (List1.Text)
End Sub

Private Sub cmdDetalladoAdj_Click()
On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsprevio

Dim lnAge As Integer

    If optImpresion(2).value = True Then
    
      If MsgBox("Este reporte puede demorar unos minutos..." & vbCrLf & "¿Desea procesar la información ?", vbOKOnly + vbQuestion, "AVISO") = vbNo Then
          Exit Sub
      End If
        'For Lnage = 1 To frmSelectAgencias.List1.ListCount
        '    If frmSelectAgencias.List1.Selected(Lnage - 1) = True Then
                         Call ExportarExcel(2)
         '                cconta = cconta + 1
                         'Exit Sub
        '    End If
        'Next Lnage
    End If
    
    
    lsCadImprimir = ""
    
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
        'ImprimeContRema
                lsCadImprimir = lsCadImprimir & loImprime.nImprimeListadoParaAdjConSIAF(Format(Me.txtFecAdjudica.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, IIf(Me.txtEstado.Text = "NO INICIADO", "0000", Me.txtNumAdjudica.Text), gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, _
                        CCur(val(Me.txtPreOro14.Text)), CCur(val(Me.txtPreOro16.Text)), CCur(val(Me.txtPreOro18.Text)), CCur(val(Me.txtPreOro21.Text)), _
                        gsNomCmac, gsNomAge, gsCodUser, True, lsmensaje, gImpresora)
                        If Trim(lsmensaje) <> "" Then
                             MsgBox lsmensaje, vbInformation, "Aviso"
                             Exit Sub
                        End If
            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Me.optImpresion(0).value = True Then
        Set loPrevio = New previo.clsprevio
            loPrevio.Show lsCadImprimir, "Listado Contratos para Remate", True
        Set loPrevio = Nothing
    Else
'        dlgGrabar.CancelError = True
'        dlgGrabar.InitDir = App.path
'        dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
'        dlgGrabar.ShowSave
'        If dlgGrabar.FileName <> "" Then
'           Open me.dl .FileName For Output As #1
'            Print #1, vBuffer
'            Close #1
'        End If
    End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    If Err.Number = 32755 Then
        MsgBox " Grabación Cancelada ", vbInformation, " Aviso "
    Else
        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
    End If


End Sub

'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Private Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
      Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox Err.Description, vbInformation, "Aviso"
  ExcelBegin = False
End Function

'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Private Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox Err.Description, vbInformation, "Aviso"
End Sub


'********************************
' Adiciona Hoja a LibroExcel
'********************************
Private Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet)
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = psHojName Then
       xlHoja1.Delete
       Exit For
    End If
Next
Set xlHoja1 = xlLibro.Worksheets.Add
xlHoja1.Name = psHojName
End Sub

Private Sub ExportarExcel(ByVal nTpoReporte As Integer)
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lnAge As Integer, cconta As Integer
Dim nFila As Long, i As Long

Dim orep As NColPRecGar, rstemp As New ADODB.Recordset, lrCredPigJoyasDet As New ADODB.Recordset
Dim lsRep As String, loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim lsArchivoN As String, lbLibroOpen As Boolean

Dim iTem As Integer


lsRep = "DETALLEADJ"


lsArchivoN = App.Path & "\Spooler\Rep" & lsRep & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
     
OleExcel.Class = "ExcelWorkSheet"



lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
'SNomFile = lsArchivoN


   If lbLibroOpen Then

    cconta = 1
    For lnAge = 1 To frmSelectAgencias.List1.ListCount

            If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                    Set orep = New NColPRecGar
            
                    Select Case nTpoReporte
                            Case 2
                                Set rstemp = orep.nListadoParaAdjConSIAF(Format(Me.txtFecAdjudica.Text, "mm/dd/yyyy"), _
                                    Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, IIf(Me.txtEstado = "NO INICIADO", "0000", Me.txtNumAdjudica), gdFecSis, _
                                    fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, _
                                    CCur(val(Me.txtPreOro14.Text)), CCur(val(Me.txtPreOro16.Text)), CCur(val(Me.txtPreOro18.Text)), CCur(val(Me.txtPreOro21.Text)), _
                                    gsNomCmac, gsNomAge, gsCodUser, True)
                    End Select
            
            
                    If rstemp.EOF Then
                        MsgBox "No se encontro información para este reporte", vbOKOnly + vbInformation, "Aviso"
                        'Exit Sub
                        GoTo Fin1
                    End If
                        If cconta = 1 Then
                           Set xlHoja1 = xlLibro.Worksheets(1)
                        Else
                           Set xlHoja1 = xlLibro.Worksheets(cconta)
                        End If
                        
                        ExcelAddHoja Format(gdFecSis, "yyyymmdd") & Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), xlLibro, xlHoja1
        
                        nFila = 1
                        
                        xlHoja1.Cells(nFila, 1) = gsNomCmac
                        nFila = 2
                        xlHoja1.Cells(nFila, 1) = gsNomAge
                        xlHoja1.Range("F2:H2").MergeCells = True
                        xlHoja1.Cells(nFila, 6) = Format(gdFecSis, "Long Date")
                         
                         'prgBar.value = 2
                             
                         
                        nFila = 3
                        xlHoja1.Cells(nFila, 1) = "LISTADO DE CONTRATOS PARA ADJUDICACION NRO " & IIf(Me.txtEstado = "NO INICIADO", "0000", Me.txtNumAdjudica) & "DEL " & Format(txtFecAdjudica, "dd/mm/yyyy") & " - Agencia " & Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2)
                                                   
                        
                        xlHoja1.Range("A1:M5").Font.Bold = True
                        
                        xlHoja1.Range("A3:M3").MergeCells = True
                        xlHoja1.Range("A3:A3").HorizontalAlignment = xlCenter
                        xlHoja1.Range("A5:M5").HorizontalAlignment = xlCenter
             
                        'xlHoja1.Range("A5:H5").AutoFilter
                        
                        nFila = 5
                        
                        xlHoja1.Cells(nFila, 1) = "ITEM "
                        xlHoja1.Cells(nFila, 2) = "CUENTA SIAF "
                        xlHoja1.Cells(nFila, 3) = "CUENTA SICMACI "
                        xlHoja1.Cells(nFila, 4) = "CLIENTE "
                        xlHoja1.Cells(nFila, 5) = "PIEZAS "
                        xlHoja1.Cells(nFila, 6) = "DESCRIPCION "
                        xlHoja1.Cells(nFila, 7) = "               "
                        xlHoja1.Cells(nFila, 8) = "               "
                        xlHoja1.Cells(nFila, 9) = "14K "
                        xlHoja1.Cells(nFila, 10) = "16K "
                        xlHoja1.Cells(nFila, 11) = "18K "
                        xlHoja1.Cells(nFila, 12) = "21K "
                        xlHoja1.Cells(nFila, 13) = "VAL REGISTRO "
                        
                        
            '                    lsCadImp = lsCadImp & ImpreFormat(lnIndice, 5, 0) & Space(1) & !CODESIAF & Space(1) & Mid(!CCTACOD, 1, 5) & "-" & Mid(!CCTACOD, 6) & Space(1) & ImpreFormat(!npiezas, 3, 0) & Space(1) _
            '                                        & lmDetalle(0) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) _
            '                                        & ImpreFormat(lnPreVenta, 8) & gPrnSaltoLinea
            '
                        i = 0
                        While Not rstemp.EOF
                            nFila = nFila + 1
                            
                            'prgBar.value = ((i) / RSTEMP.RecordCount) * 100
                            
                            i = i + 1
                            
                            xlHoja1.Cells(nFila, 1) = Format(i, "0000")
                            xlHoja1.Cells(nFila, 2) = rstemp!codesiaf
                            xlHoja1.Cells(nFila, 3) = rstemp!cCtaCod
                            xlHoja1.Cells(nFila, 4) = rstemp!cNomCliente
                            xlHoja1.Cells(nFila, 5) = rstemp!npiezas
                            xlHoja1.Cells(nFila, 9) = Format(rstemp!nK14, "#0.00")
                            xlHoja1.Cells(nFila, 10) = Format(rstemp!nK16, "#0.00")
                            xlHoja1.Cells(nFila, 11) = Format(rstemp!nK18, "#0.00")
                            xlHoja1.Cells(nFila, 12) = Format(rstemp!nK21, "#0.00")
                            xlHoja1.Cells(nFila, 13) = Format(rstemp!nAdjValRegistro, "#,##0.00")
                            
                            
                            Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
                                    Set lrCredPigJoyasDet = loMuestraContrato.dObtieneDatosCreditoPignoraticioJoyasDet(rstemp!cCtaCod)
                                Set loMuestraContrato = Nothing
                                
                                'Item = 0
                                Do While Not lrCredPigJoyasDet.EOF
                                
                                    nFila = nFila + 1
                                    
                                    'cKilataje,nItem, nPiezas, nPesoBruto, nPesoNeto, nValTasac, cDescrip
                                    
                                    'Item = Item + 1
                                    xlHoja1.Cells(nFila, 5) = Format(lrCredPigJoyasDet!nItem, "00")
                                    xlHoja1.Cells(nFila, 6) = ImpreFormat(lrCredPigJoyasDet!npiezas, 4, 0)
                                    xlHoja1.Cells(nFila, 7) = ImpreFormat(lrCredPigJoyasDet!cdescrip, 27)
                                    xlHoja1.Cells(nFila, 8) = lrCredPigJoyasDet!ckilataje & "K "
                                    xlHoja1.Cells(nFila, 9) = ImpreFormat(lrCredPigJoyasDet!nPesoBruto, 4, 2) & "Gr"
                                    xlHoja1.Cells(nFila, 10) = ImpreFormat(lrCredPigJoyasDet!npesoneto, 4, 2) & "Gr"
                                    lrCredPigJoyasDet.MoveNext
                                Loop
                                                        
                            rstemp.MoveNext
                            
                        Wend
                        
                        xlHoja1.Cells.Select
                        xlHoja1.Cells.Font.Name = "Arial"
                        xlHoja1.Cells.Font.Size = 9
                        xlHoja1.Cells.EntireColumn.AutoFit
                       
                        cconta = cconta + 1
                        
Fin1:
                  
                  End If
           Next lnAge
        
                
            'Cierro...
           OleExcel.Class = "ExcelWorkSheet"
           ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
           OleExcel.SourceDoc = lsArchivoN
            
                      
            
            OleExcel.Verb = 1
            OleExcel.Action = 1
            OleExcel.DoVerb -1
            
           ' prgBar.value = 100
            
   End If
   
   Set rstemp = Nothing
   

    
End Sub

'Permite editar los campos editables de un remate
Private Sub cmdEditar_Click()
    Call HabilitaControles(False, True, False, True, True, True, True, True, True, False, False, False, False, False)
'    Me.txtFecAdjudica.SetFocus
End Sub

'permite grabar los cambios ingresados
Private Sub cmdGrabar_Click()

Dim loGrabRem As COMNColoCPig.NCOMColPRecGar

On Error GoTo ControlError

'*** PEAC 20080515
'If Not IsDate(Me.txtFecIni.Text) Then
'    MsgBox "Ingrese una Fecha hasta Adjudicar", vbInformation, "Aviso"
'    Exit Sub
'End If

Call HabilitaControles(True, False, True, False, False, False, False, False, False, False, True, True, True, True)

'If (txtPreOroInternacional * txtTipoCambio) > 0 And IsDate(txtFecini) Then
'    cmdImpPlanPrevAdju.Enabled = True
'End If
'If (txtPreOroInternacional * txtTipoCambio) > 0 And DateDiff("d", txtFecAdjudica, Format(gdFecSis, "dd/mm/yyyy")) = 0 And IsDate(txtFecini) Then
'    cmdImpPlanPrevAdju.Enabled = True
'    cmdAdjuLote.Enabled = True
'End If

'*** PEAC 20080515
'    If Val(txtPreOro14) > 0 And Val(txtPreOro18) > 0 _
'        And Val(txtPreOro21) > 0 And IsDate(txtFecIni) Then
'        cmdImpPlanPrevAdju.Enabled = True
'    End If
    If val(txtPreOro14) > 0 And val(txtPreOro18) > 0 _
        And val(txtPreOro21) > 0 Then
        cmdImpPlanPrevAdju.Enabled = True
    End If



'*** PEAC 20080515
'    If Val(txtPreOro14) > 0 And Val(txtPreOro18) > 0 _
'        And Val(txtPreOro21) > 0 And IsDate(txtFecIni) _
'        And DateDiff("d", txtFecAdjudica, Format(gdFecSis, "dd/mm/yyyy")) = 0 And IsDate(txtFecIni) Then
'        cmdImpPlanPrevAdju.Enabled = True
'        cmdAdjuLote.Enabled = True
'    End If

    If val(txtPreOro14) > 0 And val(txtPreOro18) > 0 _
        And val(txtPreOro21) > 0 _
        And DateDiff("d", txtFecAdjudica, Format(gdFecSis, "dd/mm/yyyy")) = 0 Then
        cmdImpPlanPrevAdju.Enabled = True
        cmdAdjuLote.Enabled = True
    End If


   
Set loGrabRem = New COMNColoCPig.NCOMColPRecGar
    'Call loGrabRem.nRecGarGrabaDatosPreparaCredPignoraticio("A", Me.txtNumAdjudica.Text, gColPRecGarEstNoIniciado, Format(Me.txtFecAdjudica.Text, "mm/dd/yyyy hh:mm"), fsProcesoCadaAgencia, CCur(Val(Me.txtPreOro14.Text)), CCur(Val(Me.txtPreOro16.Text)), CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), , , Format(Me.txtFecIni.Text, "mm/dd/yyyy"), True)
    Call loGrabRem.nRecGarGrabaDatosPreparaCredPignoraticio("A", Me.txtNumAdjudica.Text, gColPRecGarEstNoIniciado, Format(Me.txtFecAdjudica.Text, "mm/dd/yyyy hh:mm"), fsProcesoCadaAgencia, CCur(val(Me.txtPreOro14.Text)), CCur(val(Me.txtPreOro16.Text)), CCur(val(Me.txtPreOro18.Text)), CCur(val(Me.txtPreOro21.Text)), , , Format(gdFecSis, "mm/dd/yyyy"), True)
    
Set loGrabRem = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdImpListAdju_Click()
On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsprevio

Dim lnAge As Integer
    lsCadImprimir = ""
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimeListadoAdjudicados(Format(Me.txtFecAdjudica.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, IIf(Me.txtEstado.Text = "NO INICIADO", "0000", Me.txtNumAdjudica), gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, _
                        CCur(val(Me.txtPreOro14.Text)), CCur(val(Me.txtPreOro16.Text)), CCur(val(Me.txtPreOro18.Text)), CCur(val(Me.txtPreOro21.Text)), _
                        gsNomCmac, gsNomAge, gsCodUser, lsmensaje, gImpresora)
                        
                        If Trim(lsmensaje) <> "" Then
                             MsgBox lsmensaje, vbInformation, "Aviso"
                             Exit Sub
                        End If
                        
            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No se hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    Set loPrevio = New previo.clsprevio
        loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", True
    Set loPrevio = Nothing

Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
'Procedimiento de impresión de listado de adjudicados
Private Sub ImprimeListAdju(vConexion As ADODB.Connection)
'    Dim vIndice As Long  'contador de Item
'    Dim vLineas As Integer
'    'Dim vPage As Integer
'    Dim vCuenta As Integer
'    Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
'    Dim vSum14 As Currency, vSum16 As Currency, vSum18 As Currency, vSum21 As Currency
'    Dim vSumValReg As Currency
'    Dim vSumDeuda As Currency
'    Dim vVerCodAnt As String, vNroAdjAnt As String
'    MousePointer = 11
'    MuestraImpresion = True
'    'vRTFImp = ""
'    sSQL = "SELECT da.cCodCta, cp.mDescLote, cp.nPiezas, da.nvalregist, cp.dFecVenc, cp.nValTasac, cp.nTasaIntVenc, " & _
'        " cp.nCostCusto, cp.cEstado, da.ndeuda, r.ccodant " & _
'        " FROM DetAdjud DA INNER JOIN CredPrenda CP ON da.ccodcta = cp.ccodcta " & _
'        " LEFT JOIN RelConNueAntPrend R ON cp.ccodcta = r.ccodnue " & _
'        " WHERE da.cnroadju = '" & List1.Text & "' AND da.cestado not in ('X')" & _
'        " ORDER BY da.cCodCta"
'    RegCredPrend.Open sSQL, vConexion, adOpenStatic, adLockOptimistic, adCmdText
'    If (RegCredPrend.BOF Or RegCredPrend.EOF) Then
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'        MuestraImpresion = False
'        MousePointer = 0
'        MsgBox " No existen Contratos Adjudicados ", vbInformation, " Aviso "
'        Exit Sub
'    Else
'        vPage = vPage + 1
'        prgList.Min = 0: vCont = 0
'        prgList.Max = RegCredPrend.RecordCount
'        If optImpresion(0).Value = True Then
'            If prgList.Max > pPrevioMax Then
'                RegCredPrend.Close
'                Set RegCredPrend = Nothing
'                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
'                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
'                MuestraImpresion = False
'                MousePointer = 0
'                Exit Sub
'            End If
'            Cabecera "ListAdju", vPage
'        Else
'            ImpreBegin True, pHojaFiMax
'            vRTFImp = ""
'            Cabecera "ListAdju", vPage
'            Print #ArcSal, ImpreCarEsp(vRTFImp);
'            vRTFImp = ""
'        End If
'        prgList.Visible = True
'        vIndice = 1:        vLineas = 7
'        v14 = 0: v16 = 0: v18 = 0: v21 = 0
'        vSum14 = 0: vSum16 = 0: vSum18 = 0: vSum21 = 0
'        vSumValReg = 0: vSumDeuda = 0
'        vCuenta = 0
'        vNroAdjAnt = "0000"
'        With RegCredPrend
'            Do While Not .EOF
'                sSQL = "SELECT * FROM Joyas WHERE ccodcta = '" & !cCodCta & "'"
'                RegJoyas.Open sSQL, vConexion, adOpenStatic, adLockOptimistic, adCmdText
'                If (RegJoyas.BOF Or RegJoyas.EOF) Then
'                    MsgBox " No existe la descripción de sus Joyas " & !cCodCta, vbExclamation, " Error del Sistema "
'                    RegJoyas.Close
'                    Set RegJoyas = Nothing
'                    RegCredPrend.Close
'                    Set RegCredPrend = Nothing
'                    Exit Sub
'                Else
'                    Do While Not RegJoyas.EOF
'                        If RegJoyas!ckilataje = "14" Then
'                            v14 = RegJoyas!nPesoOro: v14 = Round(v14, 2)
'                            vSum14 = vSum14 + v14
'                        End If
'                        If RegJoyas!ckilataje = "16" Then
'                            v16 = RegJoyas!nPesoOro: v16 = Round(v16, 2)
'                            vSum16 = vSum16 + v16
'                        End If
'                        If RegJoyas!ckilataje = "18" Then
'                            v18 = RegJoyas!nPesoOro: v18 = Round(v18, 2)
'                            vSum18 = vSum18 + v18
'                        End If
'                        If RegJoyas!ckilataje = "21" Then
'                            v21 = RegJoyas!nPesoOro: v21 = Round(v21, 2)
'                            vSum21 = vSum21 + v21
'                        End If
'                        RegJoyas.MoveNext
'                    Loop
'                    RegJoyas.Close
'                    Set RegJoyas = Nothing
'                End If
'                'Para Ver el Código Antiguo
'                'If pVerCodAnt Then vVerCodAnt = IIf(IsNull(!cCodAnt), "", !cCodAnt)
'                If optImpresion(0).Value = True Then
'                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 6, 0) & ImpreFormat(!cCodCta, 13, 1) & ImpreFormat(!npiezas, 3, 0) & _
'                        ImpreFormat(QuiebreTexto(!mDescLote, 1), 30, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) & _
'                        ImpreFormat(v14 + v16 + v18 + v21, 8) & ImpreFormat(!ndeuda, 10) & ImpreFormat(!nvalregist, 10) & gPrnSaltoLinea
'                    If pVerCodAnt And Len(vVerCodAnt) > 0 Then
'                        vRTFImp = vRTFImp & Space(7) & ImpreFormat(vVerCodAnt, 13, 0)
'                    End If
'                Else
'                    Print #ArcSal, ImpreFormat(vIndice, 6, 0) & ImpreFormat(!cCodCta, 13, 1) & ImpreFormat(!npiezas, 3, 0) & _
'                        ImpreFormat(QuiebreTexto(!mDescLote, 1), 30, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) & _
'                        ImpreFormat(v14 + v16 + v18 + v21, 8) & ImpreFormat(!ndeuda, 10) & ImpreFormat(!nvalregist, 10)
'                    If pVerCodAnt And Len(vVerCodAnt) > 0 Then
'                        Print #ArcSal, Space(7) & ImpreFormat(vVerCodAnt, 13, 0);
'                    End If
'                End If
'                vLineas = vLineas + 1
'                Dim bPasDes As Boolean
'                bPasDes = False
'                Do While vCuenta < 15
'                    vCuenta = vCuenta + 1
'                    If Len(QuiebreTexto(!mDescLote, vCuenta + 1)) > 0 Then
'                        bPasDes = True
'                        If optImpresion(0).Value = True Then
'                            If pVerCodAnt And Len(vVerCodAnt) > 0 And vCuenta = 1 Then
'                                vRTFImp = vRTFImp & Space(4) & QuiebreTexto(!mDescLote, vCuenta + 1) & gPrnSaltoLinea
'                            Else
'                                vRTFImp = vRTFImp & Space(24) & QuiebreTexto(!mDescLote, vCuenta + 1) & gPrnSaltoLinea
'                            End If
'                        Else
'                            If pVerCodAnt And Len(vVerCodAnt) > 0 And vCuenta = 1 Then
'                                Print #ArcSal, Space(4) & QuiebreTexto(!mDescLote, vCuenta + 1)
'                            Else
'                                Print #ArcSal, Space(24) & QuiebreTexto(!mDescLote, vCuenta + 1)
'                            End If
'                        End If
'                        vLineas = vLineas + 1
'                    End If
'                Loop
'                If Not bPasDes And pVerCodAnt Then
'                    If optImpresion(0).Value = True Then
'                        vRTFImp = vRTFImp & gPrnSaltoLinea
'                    Else
'                        Print #ArcSal, ""
'                    End If
'                End If
'                vSumValReg = vSumValReg + !nvalregist
'                vSumDeuda = vSumDeuda + !ndeuda
'                vCuenta = 0
'                v14 = 0: v16 = 0: v18 = 0: v21 = 0
'                vIndice = vIndice + 1
'                If vLineas >= 55 Then
'                    vPage = vPage + 1
'                    If optImpresion(0).Value = True Then
'                        vRTFImp = vRTFImp & Chr(12)
'                        Cabecera "ListAdju", vPage
'                    Else
'                        If vPage Mod 5 = 0 Then
'                            ImpreEnd
'                            ImpreBegin True, pHojaFiMax
'                        Else
'                            ImpreNewPage
'                        End If
'                        vRTFImp = ""
'                        Cabecera "ListAdju", vPage
'                        Print #ArcSal, ImpreCarEsp(vRTFImp);
'                        vRTFImp = ""
'                    End If
'                    vLineas = 7
'                End If
'                vCont = vCont + 1
'                prgList.Value = vCont
'                .MoveNext
'            Loop
'            vTot14 = vTot14 + vSum14
'            vTot16 = vTot16 + vSum16
'            vTot18 = vTot18 + vSum18
'            vTot21 = vTot21 + vSum21
'            vTotDeuda = vTotDeuda + vSumDeuda
'            vTotValReg = vTotValReg + vSumValReg
'
'            If optImpresion(0).Value = True Then
'                vRTFImp = vRTFImp & gPrnSaltoLinea
'                vRTFImp = vRTFImp & Space(48) & "Total " & ImpreFormat(vSum14, 6) & ImpreFormat(vSum16, 6) & _
'                    ImpreFormat(vSum18, 6) & ImpreFormat(vSum21, 6) & ImpreFormat(vSum14 + vSum16 + vSum18 + vSum21, 8) & ImpreFormat(vSumDeuda, 10) & ImpreFormat(vSumValReg, 10) & gPrnSaltoLinea
'                If vTotAge = 0 Then
'                    vRTFImp = vRTFImp & gPrnSaltoLinea
'                    vRTFImp = vRTFImp & Space(48) & "TOTAL " & ImpreFormat(vTot14, 6) & ImpreFormat(vTot16, 6) & _
'                        ImpreFormat(vTot18, 6) & ImpreFormat(vTot21, 6) & ImpreFormat(vTot14 + vTot16 + vTot18 + vTot21, 8) & ImpreFormat(vTotDeuda, 10) & ImpreFormat(vTotValReg, 10) & gPrnSaltoLinea
'                End If
'                vRTFImp = vRTFImp & Chr(12)
'            Else
'                Print #ArcSal, ""
'                Print #ArcSal, Space(48) & "Total " & ImpreFormat(vSum14, 6) & ImpreFormat(vSum16, 6) & _
'                    ImpreFormat(vSum18, 6) & ImpreFormat(vSum21, 6) & ImpreFormat(vSum14 + vSum16 + vSum18 + vSum21, 8) & ImpreFormat(vSumDeuda, 10) & ImpreFormat(vSumValReg, 10)
'                If vTotAge = 0 Then
'                    Print #ArcSal, ""
'                    Print #ArcSal, Space(48) & "TOTAL " & ImpreFormat(vTot14, 6) & ImpreFormat(vTot16, 6) & _
'                        ImpreFormat(vTot18, 6) & ImpreFormat(vTot21, 6) & ImpreFormat(vTot14 + vTot16 + vTot18 + vTot21, 8) & ImpreFormat(vTotDeuda, 10) & ImpreFormat(vTotValReg, 10)
'                End If
'                ImpreEnd
'            End If
'        End With
'        prgList.Visible = False
'        prgList.Value = 0
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'    End If
'    MousePointer = 0
End Sub

'Procedimiento de impresión de listado que se encuentran en ESTADO ADJUDICADO, (NO INCLUYE LOS  VENDIDOS EN SUBASTA)
Private Sub ImprimeListAdjuCons(vConexion As ADODB.Connection)
'    Dim vIndice As Long  'contador de Item
'    Dim vLineas As Integer
'    'Dim vPage As Integer
'    Dim vCuenta As Integer
'    Dim vOro14 As Currency, vOro16 As Currency, vOro18 As Currency, vOro21 As Currency
'    Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
'    Dim vSum14 As Currency, vSum16 As Currency, vSum18 As Currency, vSum21 As Currency
'    Dim vPieza As Long, vSumPieza As Long
'
'    Dim vOroValReg As Currency, vOroDeuda As Currency
'    Dim vSumValReg As Currency, vSumDeuda As Currency
'
'    Dim vVerCodAnt As String, vNroAdjAnt As String
'    MousePointer = 11
'    MuestraImpresion = True
'    'vRTFImp = ""
'    sSQL = "SELECT da.cCodCta, cp.mDescLote, cp.nPiezas, da.nvalregist, cp.dFecVenc, cp.nValTasac, cp.nTasaIntVenc, " & _
'        " cp.nCostCusto, cp.cEstado, da.ndeuda, da.cnroadju, r.ccodant, da.dFecModif " & _
'        " FROM DetAdjud DA JOIN CredPrenda CP ON da.ccodcta = cp.ccodcta " & _
'        " LEFT JOIN RelConNueAntPrend R ON cp.ccodcta = r.ccodnue " & _
'        " WHERE datediff(dd,'" & Format(pFecUltVtaBarOro, "mm/dd/yyyy") & "', da.dfecmodif) >= 0  AND cp.Cestado = '8' AND " & _
'        " datediff(dd, da.dfecmodif,'" & Format(gdFecSis, "mm/dd/yyyy") & "') >= 0 AND da.cestado not in ('X') " & _
'        " ORDER BY da.cnroadju, da.cCodCta "
'    RegCredPrend.Open sSQL, vConexion, adOpenStatic, adLockOptimistic, adCmdText
'    If (RegCredPrend.BOF Or RegCredPrend.EOF) Then
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'        MuestraImpresion = False
'        MousePointer = 0
'        MsgBox " No existen Contratos para Adjudicación ", vbInformation, " Aviso "
'        Exit Sub
'    Else
'        vPage = vPage + 1
'        prgList.Min = 0: vCont = 0
'        prgList.Max = RegCredPrend.RecordCount
'        If optImpresion(0).Value = True Then
'            If prgList.Max > pPrevioMax Then
'                RegCredPrend.Close
'                Set RegCredPrend = Nothing
'                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
'                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
'                MuestraImpresion = False
'                MousePointer = 0
'                Exit Sub
'            End If
'            Cabecera "ListAdju", vPage
'        Else
'            ImpreBegin True, pHojaFiMax
'            vRTFImp = ""
'            Cabecera "ListAdju", vPage
'            Print #ArcSal, ImpreCarEsp(vRTFImp);
'            vRTFImp = ""
'        End If
'        prgList.Visible = True
'        vIndice = 1:        vLineas = 7
'        vOro14 = 0: vOro16 = 0: vOro18 = 0: vOro21 = 0
'        v14 = 0: v16 = 0: v18 = 0: v21 = 0: vPieza = 0
'        vSum14 = 0: vSum16 = 0: vSum18 = 0: vSum21 = 0: vSumPieza = 0
'        vOroValReg = 0: vOroDeuda = 0
'        vSumValReg = 0: vSumDeuda = 0
'        vCuenta = 0
'        vNroAdjAnt = "0000"
'        With RegCredPrend
'            Do While Not .EOF
'                If Not ExisAdjVenSub(!cCodCta) Then
'                    sSQL = "SELECT * FROM Joyas WHERE ccodcta = '" & !cCodCta & "'"
'                    RegJoyas.Open sSQL, vConexion, adOpenStatic, adLockOptimistic, adCmdText
'                    If (RegJoyas.BOF Or RegJoyas.EOF) Then
'                        MsgBox " No existe la descripción de sus Joyas " & !cCodCta, vbExclamation, " Error del Sistema "
'                        RegJoyas.Close
'                        Set RegJoyas = Nothing
'                        RegCredPrend.Close
'                        Set RegCredPrend = Nothing
'                        Exit Sub
'                    Else
'                        Do While Not RegJoyas.EOF
'                            If RegJoyas!ckilataje = "14" Then
'                                v14 = RegJoyas!nPesoOro: v14 = Round(v14, 2)
'                                vSum14 = vSum14 + v14
'                            End If
'                            If RegJoyas!ckilataje = "16" Then
'                                v16 = RegJoyas!nPesoOro: v16 = Round(v16, 2)
'                                vSum16 = vSum16 + v16
'                            End If
'                            If RegJoyas!ckilataje = "18" Then
'                                v18 = RegJoyas!nPesoOro: v18 = Round(v18, 2)
'                                vSum18 = vSum18 + v18
'                            End If
'                            If RegJoyas!ckilataje = "21" Then
'                                v21 = RegJoyas!nPesoOro: v21 = Round(v21, 2)
'                                vSum21 = vSum21 + v21
'                            End If
'                            RegJoyas.MoveNext
'                        Loop
'                        RegJoyas.Close
'                        Set RegJoyas = Nothing
'                    End If
'
'                    If !cnroadju <> vNroAdjAnt Then
'                        If optImpresion(0).Value = True Then
'                            If vOro14 + vOro16 + vOro18 + vOro21 > 0 Then
'                                vRTFImp = vRTFImp & gPrnSaltoLinea
'                                vRTFImp = vRTFImp & Space(5) & "Total Adjud. " & ImpreFormat(vPieza, 6, 0) & Space(30) & ImpreFormat(vOro14, 6) & ImpreFormat(vOro16, 6) & _
'                                    ImpreFormat(vOro18, 6) & ImpreFormat(vOro21, 6) & ImpreFormat(vOro14 + vOro16 + vOro18 + vOro21, 8) & ImpreFormat(vOroDeuda, 10) & ImpreFormat(vOroValReg, 10) & gPrnSaltoLinea
'                                vOro14 = 0: vOro16 = 0: vOro18 = 0: vOro21 = 0: vOroValReg = 0: vOroDeuda = 0: vPieza = 0
'                            End If
'                            vRTFImp = vRTFImp & gPrnSaltoLinea & Chr(27) & Chr(69)
'                            vRTFImp = vRTFImp & "     Adjudicación Nro.: " & !cnroadju & "  -  Fecha : " & Format(!dFecModif, "dd/mm/yyyy") & gPrnSaltoLinea & Chr(27) & Chr(70)
'                            vRTFImp = vRTFImp & gPrnSaltoLinea
'                        Else
'                            If vOro14 + vOro16 + vOro18 + vOro21 > 0 Then
'                                Print #ArcSal, ""
'                                Print #ArcSal, Space(5) & "Total Adjud. " & ImpreFormat(vPieza, 6, 0) & Space(30) & ImpreFormat(vOro14, 6) & ImpreFormat(vOro16, 6) & _
'                                    ImpreFormat(vOro18, 6) & ImpreFormat(vOro21, 6) & ImpreFormat(vOro14 + vOro16 + vOro18 + vOro21, 8) & ImpreFormat(vOroDeuda, 10) & ImpreFormat(vOroValReg, 10)
'                                vOro14 = 0: vOro16 = 0: vOro18 = 0: vOro21 = 0: vOroValReg = 0: vOroDeuda = 0
'                            End If
'                            Print #ArcSal, ""
'                            Print #ArcSal, "     Adjudicación Nro.: " & !cnroadju
'                            Print #ArcSal, ""
'                        End If
'                        vNroAdjAnt = !cnroadju
'                    End If
'                    'Para Ver el Código Antiguo
'                    If pVerCodAnt Then vVerCodAnt = IIf(IsNull(!cCodAnt), "", !cCodAnt)
'                    If optImpresion(0).Value = True Then
'                        vRTFImp = vRTFImp & ImpreFormat(vIndice, 6, 0) & ImpreFormat(!cCodCta, 13, 1) & ImpreFormat(!npiezas, 3, 0) & _
'                            ImpreFormat(QuiebreTexto(!mDescLote, 1), 30, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) & _
'                            ImpreFormat(v14 + v16 + v18 + v21, 8) & ImpreFormat(!ndeuda, 10) & ImpreFormat(!nvalregist, 10) & gPrnSaltoLinea
'                        If pVerCodAnt And Len(vVerCodAnt) > 0 Then
'                            vRTFImp = vRTFImp & Space(7) & ImpreFormat(vVerCodAnt, 13, 0)
'                        End If
'                    Else
'                        Print #ArcSal, ImpreFormat(vIndice, 6, 0) & ImpreFormat(!cCodCta, 13, 1) & ImpreFormat(!npiezas, 3, 0) & _
'                            ImpreFormat(QuiebreTexto(!mDescLote, 1), 30, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) & _
'                            ImpreFormat(v14 + v16 + v18 + v21, 8) & ImpreFormat(!ndeuda, 10) & ImpreFormat(!nvalregist, 10)
'                        If pVerCodAnt And Len(vVerCodAnt) > 0 Then
'                            Print #ArcSal, Space(7) & ImpreFormat(vVerCodAnt, 13, 0);
'                        End If
'                    End If
'                    vLineas = vLineas + 1
'                    Dim bPasDes As Boolean
'                    bPasDes = False
'                    Do While vCuenta < 15
'                        vCuenta = vCuenta + 1
'                        If Len(QuiebreTexto(!mDescLote, vCuenta + 1)) > 0 Then
'                            bPasDes = True
'                            If optImpresion(0).Value = True Then
'                                If pVerCodAnt And Len(vVerCodAnt) > 0 And vCuenta = 1 Then
'                                    vRTFImp = vRTFImp & Space(4) & QuiebreTexto(!mDescLote, vCuenta + 1) & gPrnSaltoLinea
'                                Else
'                                    vRTFImp = vRTFImp & Space(24) & QuiebreTexto(!mDescLote, vCuenta + 1) & gPrnSaltoLinea
'                                End If
'                            Else
'                                If pVerCodAnt And Len(vVerCodAnt) > 0 And vCuenta = 1 Then
'                                    Print #ArcSal, Space(4) & QuiebreTexto(!mDescLote, vCuenta + 1)
'                                Else
'                                    Print #ArcSal, Space(24) & QuiebreTexto(!mDescLote, vCuenta + 1)
'                                End If
'                            End If
'                            vLineas = vLineas + 1
'                        End If
'                    Loop
'                    If Not bPasDes And pVerCodAnt Then
'                        If optImpresion(0).Value = True Then
'                            vRTFImp = vRTFImp & gPrnSaltoLinea
'                        Else
'                            Print #ArcSal, ""
'                        End If
'                    End If
'                    vOro14 = vOro14 + v14
'                    vOro16 = vOro16 + v16
'                    vOro18 = vOro18 + v18
'                    vOro21 = vOro21 + v21
'                    vOroValReg = vOroValReg + !nvalregist
'                    vOroDeuda = vOroDeuda + !ndeuda
'                    vPieza = vPieza + !npiezas ' ***
'
'                    vSumValReg = vSumValReg + !nvalregist
'                    vSumDeuda = vSumDeuda + !ndeuda
'                    vSumPieza = vSumPieza + !npiezas ' ***
'                    vCuenta = 0
'                    v14 = 0: v16 = 0: v18 = 0: v21 = 0
'                    vIndice = vIndice + 1
'                    If vLineas >= 55 Then
'                        vPage = vPage + 1
'                        If optImpresion(0).Value = True Then
'                            vRTFImp = vRTFImp & Chr(12)
'                            Cabecera "ListAdju", vPage
'                        Else
'                            If vPage Mod 5 = 0 Then
'                                ImpreEnd
'                                ImpreBegin True, pHojaFiMax
'                            Else
'                                ImpreNewPage
'                            End If
'                            vRTFImp = ""
'                            Cabecera "ListAdju", vPage
'                            Print #ArcSal, ImpreCarEsp(vRTFImp);
'                            vRTFImp = ""
'                        End If
'                        vLineas = 7
'                    End If
'                End If
'                vCont = vCont + 1
'                prgList.Value = vCont
'                .MoveNext
'            Loop
'            vTot14 = vTot14 + vSum14
'            vTot16 = vTot16 + vSum16
'            vTot18 = vTot18 + vSum18
'            vTot21 = vTot21 + vSum21
'            vTotDeuda = vTotDeuda + vSumDeuda
'            vTotValReg = vTotValReg + vSumValReg
'            vTotPieza = vTotPieza + vSumPieza  '***
'
'            If optImpresion(0).Value = True Then
'                vRTFImp = vRTFImp & gPrnSaltoLinea
'                vRTFImp = vRTFImp & Space(5) & "Total Adjud. " & ImpreFormat(vPieza, 6, 0) & Space(30) & ImpreFormat(vOro14, 6) & ImpreFormat(vOro16, 6) & _
'                    ImpreFormat(vOro18, 6) & ImpreFormat(vOro21, 6) & ImpreFormat(vOro14 + vOro16 + vOro18 + vOro21, 8) & ImpreFormat(vOroDeuda, 10) & ImpreFormat(vOroValReg, 10) & gPrnSaltoLinea
'
'                vRTFImp = vRTFImp & gPrnSaltoLinea
'                vRTFImp = vRTFImp & Space(5) & "Total Agencia" & ImpreFormat(vSumPieza, 6, 0) & Space(30) & ImpreFormat(vSum14, 6) & ImpreFormat(vSum16, 6) & _
'                    ImpreFormat(vSum18, 6) & ImpreFormat(vSum21, 6) & ImpreFormat(vSum14 + vSum16 + vSum18 + vSum21, 8) & ImpreFormat(vSumDeuda, 10) & ImpreFormat(vSumValReg, 10) & gPrnSaltoLinea
'                If vTotAge = 0 Then
'                    vRTFImp = vRTFImp & gPrnSaltoLinea
'                    vRTFImp = vRTFImp & Space(5) & "TOTAL GNRAL.:" & ImpreFormat(vTotPieza, 6, 0) & Space(30) & ImpreFormat(vTot14, 6) & ImpreFormat(vTot16, 6) & _
'                    ImpreFormat(vTot18, 6) & ImpreFormat(vTot21, 6) & ImpreFormat(vTot14 + vTot16 + vTot18 + vTot21, 8) & ImpreFormat(vTotDeuda, 10) & ImpreFormat(vTotValReg, 10) & gPrnSaltoLinea
'                End If
'                vRTFImp = vRTFImp & Chr(12)
'            Else
'                Print #ArcSal, ""
'                Print #ArcSal, Space(5) & "Total Adjud. " & ImpreFormat(vPieza, 6, 0) & Space(30) & ImpreFormat(vOro14, 6) & ImpreFormat(vOro16, 6) & _
'                    ImpreFormat(vOro18, 6) & ImpreFormat(vOro21, 6) & ImpreFormat(vOro14 + vOro16 + vOro18 + vOro21, 8) & ImpreFormat(vOroDeuda, 10) & ImpreFormat(vOroValReg, 10)
'
'                Print #ArcSal, ""
'                Print #ArcSal, Space(5) & "Total Agencia" & ImpreFormat(vSumPieza, 6, 0) & Space(30) & ImpreFormat(vSum14, 6) & ImpreFormat(vSum16, 6) & _
'                    ImpreFormat(vSum18, 6) & ImpreFormat(vSum21, 6) & ImpreFormat(vSum14 + vSum16 + vSum18 + vSum21, 8) & ImpreFormat(vSumDeuda, 10) & ImpreFormat(vSumValReg, 10)
'                If vTotAge = 0 Then
'                    Print #ArcSal, ""
'                    Print #ArcSal, Space(5) & "TOTAL GNRAL.:" & ImpreFormat(vTotPieza, 6, 0) & Space(30) & ImpreFormat(vTot14, 6) & ImpreFormat(vTot16, 6) & _
'                        ImpreFormat(vTot18, 6) & ImpreFormat(vTot21, 6) & ImpreFormat(vTot14 + vTot16 + vTot18 + vTot21, 8) & ImpreFormat(vTotDeuda, 10) & ImpreFormat(vTotValReg, 10)
'                End If
'                ImpreEnd
'            End If
'        End With
'        prgList.Visible = False
'        prgList.Value = 0
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'    End If
'    MousePointer = 0
End Sub

Private Sub cmdImpListAdjuMes_Click()
'On Error GoTo ControlError
'    Dim x As Integer
'    Dim pPaso  As Boolean
'    pPaso = False
'    vRTFImp = ""
'    vPage = 0
'    vTot14 = 0: vTot16 = 0: vTot18 = 0: vTot21 = 0
'    vTotDeuda = 0: vTotValReg = 0: vTotPieza = 0
'    vTotAge = 0
'    For x = 1 To frmPigAgencias.List1.ListCount
'        If frmPigAgencias.List1.Selected(x - 1) = True Then
'            vTotAge = vTotAge + 1
'        End If
'    Next x
'    For x = 1 To frmPigAgencias.List1.ListCount
'        If frmPigAgencias.List1.Selected(x - 1) = True Then
'            vTotAge = vTotAge - 1
'            If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
'                vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAge & "'")
'                Call ImprimeListAdjuMes(dbCmact)
'                pPaso = True
'            Else
'                If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
'                    vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
'                    Call ImprimeListAdjuMes(dbCmactN)
'                    pPaso = True
'                End If
'                CierraConeccion
'            End If
'        End If
'    Next x
'    If pPaso And Len(vRTFImp) > 10 Then
'        If MuestraImpresion And optImpresion(0).Value = True Then
'            rtfImp.Text = vRTFImp
'            frmPrevio.Previo rtfImp, "Listado de Joyas Adjudicadas en el Mes de " & Format(txtFecMes.Text, "mmmm"), True, 66
'        End If
'    End If
'Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
'        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Procedimiento de impresión de listado de adjudicados por mes - incluye los adjudicados y vendidos en subasta
Private Sub ImprimeListAdjuMes(vConexion As ADODB.Connection)
'    Dim vIndice As Long  'contador de Item
'    Dim vLineas As Integer
'    Dim vCuenta As Integer
'    Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency, vPieza As Long
'    Dim vSum14 As Currency, vSum16 As Currency, vSum18 As Currency, vSum21 As Currency, vSumPieza As Long
'    Dim vSumValReg As Currency
'    Dim vSumDeuda As Currency
'    Dim vVerCodAnt As String, vNroAdjAnt As String
'    MousePointer = 11
'    MuestraImpresion = True
'    'vRTFImp = ""
'
'    sSQL = "SELECT da.cCodCta, cp.mDescLote, cp.nPiezas, da.nvalregist, cp.dFecVenc, cp.nValTasac, cp.nTasaIntVenc, " & _
'        " cp.nCostCusto, cp.cEstado, da.ndeuda, da.cnroadju, r.ccodant " & _
'        " FROM DetAdjud DA INNER JOIN CredPrenda CP ON da.ccodcta = cp.ccodcta " & _
'        " LEFT JOIN RelConNueAntPrend R ON cp.ccodcta = r.ccodnue " & _
'        " WHERE datediff(mm,da.dFecmodif , '" & Format(txtFecMes.Text, "mm/dd/yyyy") & "') = 0 AND da.cestado not in ('X') " & _
'        " ORDER BY da.cnroadju , da.cCodCta"
'
'    RegCredPrend.CursorLocation = adUseClient
'    RegCredPrend.Open sSQL, vConexion, adOpenStatic, adLockOptimistic, adCmdText
'    Set RegCredPrend.ActiveConnection = Nothing
'
'    If (RegCredPrend.BOF Or RegCredPrend.EOF) Then
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'        MuestraImpresion = False
'        MousePointer = 0
'        MsgBox " No existen Contratos Adjudicados en este Mes ", vbInformation, " Aviso "
'        Exit Sub
'    Else
'        vPage = vPage + 1
'        prgList.Min = 0: vCont = 0
'        prgList.Max = RegCredPrend.RecordCount
'        If optImpresion(0).Value = True Then
'            If prgList.Max > pPrevioMax Then
'                RegCredPrend.Close
'                Set RegCredPrend = Nothing
'                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
'                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
'                MuestraImpresion = False
'                MousePointer = 0
'                Exit Sub
'            End If
'            Cabecera "ListAdjuMes", vPage
'        Else
'            ImpreBegin True, pHojaFiMax
'            vRTFImp = ""
'            Cabecera "ListAdjuMes", vPage
'            Print #ArcSal, ImpreCarEsp(vRTFImp);
'            vRTFImp = ""
'        End If
'        prgList.Visible = True
'        vIndice = 1:        vLineas = 7
'        v14 = 0: v16 = 0: v18 = 0: v21 = 0
'        vSum14 = 0: vSum16 = 0: vSum18 = 0: vSum21 = 0: vSumPieza = 0
'        vSumValReg = 0: vSumDeuda = 0
'        vCuenta = 0
'        vNroAdjAnt = "0000"
'        With RegCredPrend
'            Do While Not .EOF
'                sSQL = "SELECT * FROM Joyas WHERE ccodcta = '" & !cCodCta & "'"
'                RegJoyas.Open sSQL, vConexion, adOpenStatic, adLockOptimistic, adCmdText
'                If (RegJoyas.BOF Or RegJoyas.EOF) Then
'                    MsgBox " No existe la descripción de sus Joyas " & !cCodCta, vbExclamation, " Error del Sistema "
'                    RegJoyas.Close
'                    Set RegJoyas = Nothing
'                    RegCredPrend.Close
'                    Set RegCredPrend = Nothing
'                    Exit Sub
'                Else
'                    Do While Not RegJoyas.EOF
'                        If RegJoyas!ckilataje = "14" Then
'                            v14 = RegJoyas!nPesoOro: v14 = Round(v14, 2)
'                            vSum14 = vSum14 + v14
'                        End If
'                        If RegJoyas!ckilataje = "16" Then
'                            v16 = RegJoyas!nPesoOro: v16 = Round(v16, 2)
'                            vSum16 = vSum16 + v16
'                        End If
'                        If RegJoyas!ckilataje = "18" Then
'                            v18 = RegJoyas!nPesoOro: v18 = Round(v18, 2)
'                            vSum18 = vSum18 + v18
'                        End If
'                        If RegJoyas!ckilataje = "21" Then
'                            v21 = RegJoyas!nPesoOro: v21 = Round(v21, 2)
'                            vSum21 = vSum21 + v21
'                        End If
'                        RegJoyas.MoveNext
'                    Loop
'                    RegJoyas.Close
'                    Set RegJoyas = Nothing
'                End If
'                If !cnroadju <> vNroAdjAnt Then
'                    If optImpresion(0).Value = True Then
'                        vRTFImp = vRTFImp & gPrnSaltoLinea
'                        vRTFImp = vRTFImp & "     Adjudicación Nro.: " & !cnroadju & gPrnSaltoLinea
'                        vRTFImp = vRTFImp & gPrnSaltoLinea
'                    Else
'                        Print #ArcSal, ""
'                        Print #ArcSal, "     Adjudicación Nro.: " & !cnroadju
'                        Print #ArcSal, ""
'                    End If
'                    vNroAdjAnt = !cnroadju
'                End If
'                'Para Ver el Código Antiguo
'                If pVerCodAnt Then vVerCodAnt = IIf(IsNull(!cCodAnt), "", !cCodAnt)
'                If optImpresion(0).Value = True Then
'                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 6, 0) & ImpreFormat(!cCodCta, 13, 1) & ImpreFormat(!npiezas, 3, 0) & _
'                        ImpreFormat(QuiebreTexto(!mDescLote, 1), 30, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) & _
'                        ImpreFormat(v14 + v16 + v18 + v21, 8) & ImpreFormat(!ndeuda, 10) & ImpreFormat(!nvalregist, 10) & gPrnSaltoLinea
'                    If pVerCodAnt And Len(vVerCodAnt) > 0 Then
'                        vRTFImp = vRTFImp & Space(7) & ImpreFormat(vVerCodAnt, 13, 0)
'                    End If
'                Else
'                    Print #ArcSal, ImpreFormat(vIndice, 6, 0) & ImpreFormat(!cCodCta, 13, 1) & ImpreFormat(!npiezas, 3, 0) & _
'                        ImpreFormat(QuiebreTexto(!mDescLote, 1), 30, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) & _
'                        ImpreFormat(v14 + v16 + v18 + v21, 8) & ImpreFormat(!ndeuda, 10) & ImpreFormat(!nvalregist, 10)
'                    If pVerCodAnt And Len(vVerCodAnt) > 0 Then
'                        Print #ArcSal, Space(7) & ImpreFormat(vVerCodAnt, 13, 0);
'                    End If
'                End If
'                vLineas = vLineas + 1
'                Dim bPasDes As Boolean
'                bPasDes = False
'                Do While vCuenta < 15
'                    vCuenta = vCuenta + 1
'                    If Len(QuiebreTexto(!mDescLote, vCuenta + 1)) > 0 Then
'                        bPasDes = True
'                        If optImpresion(0).Value = True Then
'                            If pVerCodAnt And Len(vVerCodAnt) > 0 And vCuenta = 1 Then
'                                vRTFImp = vRTFImp & Space(4) & QuiebreTexto(!mDescLote, vCuenta + 1) & gPrnSaltoLinea
'                            Else
'                                vRTFImp = vRTFImp & Space(24) & QuiebreTexto(!mDescLote, vCuenta + 1) & gPrnSaltoLinea
'                            End If
'                        Else
'                            If pVerCodAnt And Len(vVerCodAnt) > 0 And vCuenta = 1 Then
'                                Print #ArcSal, Space(4) & QuiebreTexto(!mDescLote, vCuenta + 1)
'                            Else
'                                Print #ArcSal, Space(24) & QuiebreTexto(!mDescLote, vCuenta + 1)
'                            End If
'                        End If
'                        vLineas = vLineas + 1
'                    End If
'                Loop
'                If Not bPasDes And pVerCodAnt Then
'                    If optImpresion(0).Value = True Then
'                        vRTFImp = vRTFImp & gPrnSaltoLinea
'                    Else
'                        Print #ArcSal, ""
'                    End If
'                End If
'                vSumValReg = vSumValReg + !nvalregist
'                vSumDeuda = vSumDeuda + !ndeuda
'                vSumPieza = vSumPieza + !npiezas
'                vCuenta = 0
'                v14 = 0: v16 = 0: v18 = 0: v21 = 0
'                vIndice = vIndice + 1
'                If vLineas >= 55 Then
'                    vPage = vPage + 1
'                    If optImpresion(0).Value = True Then
'                        vRTFImp = vRTFImp & Chr(12)
'                        Cabecera "ListAdjuMes", vPage
'                    Else
'                        If vPage Mod 5 = 0 Then
'                            ImpreEnd
'                            ImpreBegin True, pHojaFiMax
'                        Else
'                            ImpreNewPage
'                        End If
'                        vRTFImp = ""
'                        Cabecera "ListAdjuMes", vPage
'                        Print #ArcSal, ImpreCarEsp(vRTFImp);
'                        vRTFImp = ""
'                    End If
'                    vLineas = 7
'                End If
'                vCont = vCont + 1
'                prgList.Value = vCont
'                .MoveNext
'            Loop
'            vTot14 = vTot14 + vSum14
'            vTot16 = vTot16 + vSum16
'            vTot18 = vTot18 + vSum18
'            vTot21 = vTot21 + vSum21
'            vTotDeuda = vTotDeuda + vSumDeuda
'            vTotValReg = vTotValReg + vSumValReg
'            vTotPieza = vTotPieza + vSumPieza
'
'            If optImpresion(0).Value = True Then
'                vRTFImp = vRTFImp & gPrnSaltoLinea
'                vRTFImp = vRTFImp & Space(5) & "Total  " & Space(5) & ImpreFormat(vSumPieza, 6, 0) & Space(31) & ImpreFormat(vSum14, 6) & ImpreFormat(vSum16, 6) & _
'                    ImpreFormat(vSum18, 6) & ImpreFormat(vSum21, 6) & ImpreFormat(vSum14 + vSum16 + vSum18 + vSum21, 8) & ImpreFormat(vSumDeuda, 10) & ImpreFormat(vSumValReg, 10) & gPrnSaltoLinea
'                If vTotAge = 0 Then
'                    vRTFImp = vRTFImp & gPrnSaltoLinea
'                    vRTFImp = vRTFImp & Space(5) & "TOTAL  " & Space(5) & ImpreFormat(vTotPieza, 6, 0) & Space(31) & ImpreFormat(vTot14, 6) & ImpreFormat(vTot16, 6) & _
'                        ImpreFormat(vTot18, 6) & ImpreFormat(vTot21, 6) & ImpreFormat(vTot14 + vTot16 + vTot18 + vTot21, 8) & ImpreFormat(vTotDeuda, 10) & ImpreFormat(vTotValReg, 10) & gPrnSaltoLinea
'                End If
'                vRTFImp = vRTFImp & Chr(12)
'            Else
'                Print #ArcSal, ""
'                Print #ArcSal, Space(5) & "Total  " & Space(5) & ImpreFormat(vSumPieza, 6, 0) & Space(31) & ImpreFormat(vSum14, 6) & ImpreFormat(vSum16, 6) & _
'                    ImpreFormat(vSum18, 6) & ImpreFormat(vSum21, 6) & ImpreFormat(vSum14 + vSum16 + vSum18 + vSum21, 8) & ImpreFormat(vSumDeuda, 10) & ImpreFormat(vSumValReg, 10)
'                If vTotAge = 0 Then
'                    Print #ArcSal, ""
'                    Print #ArcSal, Space(5) & "TOTAL  " & Space(5) & ImpreFormat(vTotPieza, 6, 0) & Space(31) & ImpreFormat(vTot14, 6) & ImpreFormat(vTot16, 6) & _
'                        ImpreFormat(vTot18, 6) & ImpreFormat(vTot21, 6) & ImpreFormat(vTot14 + vTot16 + vTot18 + vTot21, 8) & ImpreFormat(vTotDeuda, 10) & ImpreFormat(vTotValReg, 10)
'                End If
'                ImpreEnd
'            End If
'        End With
'        prgList.Visible = False
'        prgList.Value = 0
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'    End If
'    MousePointer = 0
End Sub

Private Sub cmdPlanPreviaAnt_Click()

On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsprevio

Dim lnAge As Integer
    lsCadImprimir = ""
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimePlanillaParaAdjudicacionConSIAF(Format(Me.txtFecAdjudica.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, fnNroRematesAdjudicacion, gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, _
                        CCur(val(Me.txtPreOro14.Text)), CCur(val(Me.txtPreOro16.Text)), CCur(val(Me.txtPreOro18.Text)), CCur(val(Me.txtPreOro21.Text)), _
                        gsNomCmac, gsNomAge, gsCodUser, Me.txtNumAdjudica.Text, gdFecSis, lsmensaje, gImpresora)
                If Trim(lsmensaje) <> "" Then
                    MsgBox lsmensaje, vbInformation, "Aviso"
                    Exit Sub
                End If
            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    Set loPrevio = New previo.clsprevio
        loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", True
    Set loPrevio = Nothing

Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

'Permite salir del formulario actual
Private Sub cmdsalir_Click()
    Unload frmSelectAgencias
    Unload Me
End Sub

Private Sub cmPlanAnt_Click()
On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsprevio

Dim lnAge As Integer
    lsCadImprimir = ""
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimeListadoAdjudicadosConSIAF(Format(Me.txtFecAdjudica.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, IIf(Me.txtEstado.Text = "NO INICIADO", "0000", Me.txtNumAdjudica.Text), gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, _
                        CCur(val(Me.txtPreOro14.Text)), CCur(val(Me.txtPreOro16.Text)), CCur(val(Me.txtPreOro18.Text)), CCur(val(Me.txtPreOro21.Text)), _
                        gsNomCmac, gsNomAge, gsCodUser, lsmensaje, gImpresora)
                        If Trim(lsmensaje) <> "" Then
                             MsgBox lsmensaje, vbInformation, "Aviso"
                             Exit Sub
                        End If
            Set loImprime = Nothing
                
        End If
    Next lnAge
   
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No se hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set loPrevio = New previo.clsprevio
        loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", True
    Set loPrevio = Nothing

Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

'Permite inicializar el formulario actual
Private Sub Form_Load()
    CargaParametros
    CargaAdjudicaciones
    cmdDetalladoAdj.Enabled = True
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub
Private Sub CargaAdjudicaciones()

Dim loDatos As COMNColoCPig.NCOMColPRecGar
Dim lrAdjudic As ADODB.Recordset

Set lrAdjudic = New ADODB.Recordset
Set loDatos = New COMNColoCPig.NCOMColPRecGar
    Set lrAdjudic = loDatos.nObtieneListadoProcesosRG("A", fsProcesoCadaAgencia, False)
Set loDatos = Nothing
'Mostrar Datos
If lrAdjudic Is Nothing Then
    MsgBox " No existe ninguna Adjudicación ", vbCritical, " Error de Sistema "
    cmdEditar.Enabled = False
    cmdImpPlanPrevAdju.Enabled = False
    cmdAdjuLote.Enabled = False
    cmdImpListAdju.Enabled = False
    Exit Sub
Else
    List1.Clear
    Do While Not lrAdjudic.EOF
        List1.AddItem lrAdjudic!cNroProceso
        lrAdjudic.MoveNext
    Loop
    lrAdjudic.Close
    Set lrAdjudic = Nothing
    If List1.ListCount > 1 Then List1.ListIndex = 0
End If

End Sub

Private Sub CargaDatosAdjudicacion(ByVal psNroAdju As String)

Dim loDatos As COMNColoCPig.NCOMColPRecGar
Dim lrdatosAdj As ADODB.Recordset
Dim lsmensaje As String

txtPreOro14 = Format(0, "#0.00")
txtPreOro16 = Format(0, "#0.00")
txtPreOro18 = Format(0, "#0.00")
txtPreOro21 = Format(0, "#0.00")

Set lrdatosAdj = New ADODB.Recordset
Set loDatos = New COMNColoCPig.NCOMColPRecGar
    Set lrdatosAdj = loDatos.nObtieneDatosProcesoRGCredPig("A", psNroAdju, fsProcesoCadaAgencia, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
Set loDatos = Nothing
If lrdatosAdj Is Nothing Then Exit Sub

'Mostrar Datos
txtNumAdjudica = lrdatosAdj!cNroProceso
txtFecAdjudica = Format(lrdatosAdj!dProceso, "dd/mm/yyyy")
txtHorAdjudica = Format(lrdatosAdj!dProceso, "hh:mm")

'*** PEAC 20080714
txtFecCorte = Format(lrdatosAdj!cFecCorte, "dd/mm/yyyy")

txtFecIni = IIf(IsNull(lrdatosAdj!dParaAdjudicar), "__/__/____", Format(lrdatosAdj!dParaAdjudicar, "dd/mm/yyyy"))
'** Antes
'txtPreOroInternacional = Format(lrDatosAdj!nPrecioOro, "#0.000")
'txtTipoCambio = Format(lrDatosAdj!nTipoCambio, "#0.000")
'txtPreOro14 = Format(14 * PrecioIntern(Val(txtPreOroInternacional), Val(txtTipoCambio)), "#0.00")
'txtPreOro16 = Format(16 * PrecioIntern(Val(txtPreOroInternacional), Val(txtTipoCambio)), "#0.00")
'txtPreOro18 = Format(18 * PrecioIntern(Val(txtPreOroInternacional), Val(txtTipoCambio)), "#0.00")
'txtPreOro21 = Format(21 * PrecioIntern(Val(txtPreOroInternacional), Val(txtTipoCambio)), "#0.00")
'** Ahora
txtPreOro14 = Format(lrdatosAdj!nPrecioK14, "#0.00")
txtPreOro16 = Format(lrdatosAdj!nPrecioK16, "#0.00")
txtPreOro18 = Format(lrdatosAdj!nPrecioK18, "#0.00")
txtPreOro21 = Format(lrdatosAdj!nPrecioK21, "#0.00")

txtEstado = Switch(lrdatosAdj!nRGEstado = gColPRecGarEstNoIniciado, "NO INICIADO", lrdatosAdj!nRGEstado = gColPRecGarEstIniciado, "INICIADO", lrdatosAdj!nRGEstado = gColPRecGarEstTerminado, "FINALIZADO")

If lrdatosAdj!nRGEstado = gColPRecGarEstNoIniciado Then ' "NO INICIADO"
    cmdEditar.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    'cmdImpPlanPrevAdju.Enabled = IIf(txtPreOroInternacional * txtTipoCambio > 0 And IsDate(txtFecIni), True, False)
    'cmdAdjuLote.Enabled = IIf(txtPreOroInternacional * txtTipoCambio > 0 And DateDiff("d", txtFecAdjudica, Format(gdFecSis, "dd/mm/yyyy")) = 0 And IsDate(txtFecIni), True, False)
    If val(txtPreOro14) > 0 And val(txtPreOro18) > 0 _
        And val(txtPreOro21) > 0 Then
        cmdImpPlanPrevAdju.Enabled = True
        Me.cmdPlanPreviaAnt.Enabled = True
        cmdAdjuLote.Enabled = True
    Else
        cmdImpPlanPrevAdju.Enabled = False
        Me.cmdPlanPreviaAnt.Enabled = False
        cmdAdjuLote.Enabled = False
    End If
    cmdImpListAdju.Enabled = True
    Me.cmPlanAnt.Enabled = True
    cmdImpListAdju.Caption = "Listado de Joyas Adjudicadas (Consolidado)"
ElseIf lrdatosAdj!nRGEstado = gColPRecGarEstTerminado Then  ' "FINALIZADO"
    cmdEditar.Enabled = False
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    cmdImpPlanPrevAdju.Enabled = True
    cmdPlanPreviaAnt.Enabled = True
    cmdAdjuLote.Enabled = False
    cmdImpListAdju.Enabled = True
    Me.cmPlanAnt.Enabled = True
    cmdImpListAdju.Caption = "Listado de Joyas de Adjudicación Nro.: " & List1.Text
End If

Set lrdatosAdj = Nothing

End Sub

Private Sub List1_Click()
    CargaDatosAdjudicacion (List1.Text)
End Sub

'Valida el campo txtFecAdjudica
Private Sub txtFecAdjudica_GotFocus()
fEnfoque txtFecAdjudica
End Sub
Private Sub txtFecAdjudica_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtHorAdjudica.SetFocus
End If
End Sub
Private Sub txtFecAdjudica_LostFocus()
If Not ValFecha(txtFecAdjudica) Then
    txtFecAdjudica.SetFocus
ElseIf DateDiff("d", txtFecAdjudica, gdFecSis) > 0 Then
    MsgBox " Fecha no debe ser anterior a la fecha actual", vbInformation, " Aviso "
    txtFecAdjudica.SetFocus
End If
'VeriDatRem
End Sub
Private Sub txtFecAdjudica_Validate(Cancel As Boolean)
If Not ValFecha(txtFecAdjudica) Then
    Cancel = True
ElseIf DateDiff("d", txtFecAdjudica, gdFecSis) > 0 Then
    MsgBox " Fecha no debe ser anterior a la fecha actual", vbInformation, " Aviso "
    Cancel = True
End If
End Sub

Private Sub txtFecIni_GotFocus()
fEnfoque txtFecIni
End Sub
Private Sub TxtFecIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub
Private Sub txtFecIni_LostFocus()
If Not ValFecha(txtFecIni) Then
    txtFecIni.SetFocus
ElseIf DateDiff("d", txtFecAdjudica, txtFecIni) > 0 Then
    MsgBox " Fecha no debe ser anterior a la fecha de adjudicación", vbInformation, " Aviso "
    txtFecIni.SetFocus
End If
End Sub

Private Sub txtFecIni_Validate(Cancel As Boolean)
If Not ValFecha(txtFecIni) Then
    Cancel = True
ElseIf DateDiff("d", txtFecAdjudica, txtFecIni) > 0 Then
    MsgBox " Fecha no debe ser anterior a la fecha de adjudicación", vbInformation, " Aviso "
    Cancel = True
End If
End Sub

Private Sub txtFecMes_GotFocus()
fEnfoque txtFecMes
End Sub
Private Sub txtFecMes_LostFocus()
If Not ValFecha(txtFecMes) Then
    txtFecMes.SetFocus
End If
End Sub
Private Sub txtFecMes_Validate(Cancel As Boolean)
If Not ValFecha(txtFecMes) Then
    Cancel = True
End If
End Sub

'Valida el campo txtHorAdjudica
Private Sub txtHorAdjudica_GotFocus()
fEnfoque txtHorAdjudica
End Sub
Private Sub txtHorAdjudica_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'txtPreOroInternacional.SetFocus
    txtPreOro14.SetFocus
End If
End Sub
Private Sub txtHorAdjudica_LostFocus()
If Not ValidaHora(txtHorAdjudica) Then
    txtHorAdjudica.SetFocus
End If
'VeriDatRem
End Sub

Private Sub txtPreOro14_GotFocus()
fEnfoque txtPreOro14
End Sub
Private Sub txtPreOro14_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro14, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro14 = Format(txtPreOro14, "#0.00")
    txtPreOro16.SetFocus
End If
End Sub
Private Sub txtPreOro14_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtpreoro16
Private Sub txtPreOro16_GotFocus()
fEnfoque txtPreOro16
End Sub
Private Sub txtPreOro16_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro16, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro16 = Format(txtPreOro16, "#0.00")
    txtPreOro18.SetFocus
End If
End Sub
Private Sub txtPreOro16_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtpreoro18
Private Sub txtPreOro18_GotFocus()
fEnfoque txtPreOro18
End Sub
Private Sub txtPreOro18_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro18, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro18 = Format(txtPreOro18, "#0.00")
    txtPreOro21.SetFocus
End If
End Sub
Private Sub txtPreOro18_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtpreoro21
Private Sub txtPreOro21_GotFocus()
fEnfoque txtPreOro21
End Sub
Private Sub txtPreOro21_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro21, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro21 = Format(txtPreOro21, "#0.00")
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
End If
End Sub
Private Sub txtPreOro21_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtPreOroInternacional
Private Sub txtPreOroInternacional_GotFocus()
fEnfoque txtPreOroInternacional
End Sub
Private Sub txtPreOroInternacional_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOroInternacional, KeyAscii)
If KeyAscii = 13 Then
    txtPreOroInternacional = Format(txtPreOroInternacional, "#0.00")
    txtTipoCambio.SetFocus
End If
End Sub
Private Sub txtPreOroInternacional_LostFocus()
    txtPreOroInternacional = Format(txtPreOroInternacional, "#0.00")
    VeriPreOro
End Sub

'Valida el campo txtTipoCambio
Private Sub txtTipoCambio_GotFocus()
fEnfoque txtTipoCambio
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTipoCambio, KeyAscii, 7, 3)
If KeyAscii = 13 Then
    txtTipoCambio = Format(txtTipoCambio, "#0.000")
    txtFecIni.SetFocus
    'cmdGrabar.Enabled = True
    'cmdGrabar.SetFocus
End If
End Sub
Private Sub txtTipoCambio_LostFocus()
    txtTipoCambio = Format(txtTipoCambio, "#0.000")
    VeriPreOro
End Sub

Function PrecioIntern(nPrecioOro As Double, nTipoCambio As Double) As Double
    PrecioIntern = Round((((nPrecioOro * nTipoCambio) / 31.1) / 24), 2)
End Function

' Permite activar la opción de procesar solo cuando están ingresados los campos
' fecha y hora de remate
Public Sub VeriDatRem()
    If Len(txtNumAdjudica) > 0 And Len(txtFecAdjudica) = 10 And Len(txtHorAdjudica) = 5 Then
    End If
End Sub

' Permite activar la opción de grabar solo cuando están ingresados los campos
' del precio del oro
'Private Sub VeriPreOro()
'    If Val(txtPreOroInternacional) > 0 And Val(txtTipoCambio) > 0 Then
'        txtPreOro14 = Format(14 * PrecioIntern(Val(txtPreOroInternacional), Val(txtTipoCambio)), "#0.00")
'        txtPreOro16 = Format(16 * PrecioIntern(Val(txtPreOroInternacional), Val(txtTipoCambio)), "#0.00")
'        txtPreOro18 = Format(18 * PrecioIntern(Val(txtPreOroInternacional), Val(txtTipoCambio)), "#0.00")
'        txtPreOro21 = Format(21 * PrecioIntern(Val(txtPreOroInternacional), Val(txtTipoCambio)), "#0.00")
'        cmdGrabar.Enabled = True
'    End If
'End Sub

'Imprime Prendas que serán Adjudicadas
Private Sub cmdImpPlanPrevAdju_Click()

On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsprevio

Dim lnAge As Integer
    lsCadImprimir = ""
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimePlanillaParaAdjudicacion(Format(Me.txtFecAdjudica.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, fnNroRematesAdjudicacion, gdFecSis, _
                        fnTasaCustodiaVencida, fnTasaImpuesto, fnTasaPreparacionRemate, _
                        CCur(val(Me.txtPreOro14.Text)), CCur(val(Me.txtPreOro16.Text)), CCur(val(Me.txtPreOro18.Text)), CCur(val(Me.txtPreOro21.Text)), _
                        gsNomCmac, gsNomAge, gsCodUser, Me.txtNumAdjudica.Text, gdFecSis, lsmensaje, gImpresora, Me.txtFecCorte.Text)
                        If Trim(lsmensaje) <> "" Then
                             MsgBox lsmensaje, vbInformation, "Aviso"
                             Exit Sub
                        End If
            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Len(Trim(lsCadImprimir)) = 0 Then
        MsgBox "No hay datos para mostrar en el reporte", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set loPrevio = New previo.clsprevio
        loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", True
    Set loPrevio = Nothing

Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
'Procedimiento de impresión de las planillas
Private Sub ImprimePlanPrevAdju(vConexion As ADODB.Connection)
'Dim RegPerCta As New ADODB.Recordset
'    Dim vIndice As Long  'contador de Item
'    Dim vLineas As Integer
'    'Dim vPage As Integer
'    Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
'    Dim vPreBas As Currency, vPreDeu As Currency, vPreVen As Currency
'    Dim vGramos As Currency
'    Dim vNombre As String * 28
'    Dim vSumv14 As Currency, vSumv16 As Currency, vSumv18 As Currency, vSumv21 As Currency
'    Dim vSumGramos As Currency, vSumTasaci As Currency, vSumPresta As Currency
'    Dim vSumPreDeu As Currency, vSumPreBas As Currency, vSumPreVen As Currency
'    Dim vVerCodAnt As String
'    MousePointer = 11
'    MuestraImpresion = True
'    'vRTFImp = ""
'    sSQL = "SELECT cp.cCodCta,cp.dfecPres, cp.nValTasac, cp.nPrestamo, cp.nTasaIntVenc, " & _
'        " cp.nSaldoCap, cp.dfecvenc, cp.cEstado, cp.nCostCusto, cp.nImpuesto, r.ccodant " & _
'        " FROM CredPrenda CP LEFT JOIN RelConNueAntPrend R ON cp.ccodcta = r.ccodnue " & _
'        " WHERE cp.cestado in ('1','4','6','7') and " & _
'        " DATEDIFF (dd,  cp.dFecVenc, '" & Format(txtFecIni, "mm/dd/yyyy") & "') >= 0 "
'    RegCredPrend.Open sSQL, vConexion, adOpenStatic, adLockOptimistic, adCmdText
'    If (RegCredPrend.BOF Or RegCredPrend.EOF) Then
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'        MsgBox " No existen Contratos para Adjudicación ", vbInformation, " Aviso "
'        MuestraImpresion = False
'        MousePointer = 0
'        Exit Sub
'    Else
'        vPage = 1
'        prgList.Min = 0: vCont = 0
'        prgList.Max = RegCredPrend.RecordCount
'        If optImpresion(0).Value = True Then
'            If prgList.Max > pPrevioMax Then
'                RegCredPrend.Close
'                Set RegCredPrend = Nothing
'                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
'                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
'                MuestraImpresion = False
'                MousePointer = 0
'                Exit Sub
'            End If
'            Cabecera "PlanPrAd", vPage, False
'        Else
'            ImpreBegin True, pHojaFiMax
'            vRTFImp = ""
'            Cabecera "PlanPrAd", vPage, False
'            Print #ArcSal, ImpreCarEsp(vRTFImp);
'            vRTFImp = ""
'        End If
'        prgList.Visible = True
'        vIndice = 1:        vLineas = 7
'        vPreBas = 0: vPreDeu = 0: vPreVen = 0: vGramos = 0
'        v14 = 0: v16 = 0: v18 = 0: v21 = 0
'        vSumv14 = 0:        vSumv16 = 0
'        vSumv18 = 0:        vSumv21 = 0
'        vSumGramos = 0:        vSumTasaci = 0
'        vSumPresta = 0:        vSumPreDeu = 0
'        vSumPreBas = 0:        vSumPreVen = 0
'        With RegCredPrend
'            Do While Not .EOF
'                vNombre = Left(PstaNombre(ClienteNombre(!cCodCta, vConexion), False), 30)
'                sSQL = "SELECT * FROM Joyas WHERE ccodcta = '" & !cCodCta & "'"
'                RegJoyas.Open sSQL, vConexion, adOpenStatic, adLockOptimistic, adCmdText
'                If (RegJoyas.BOF Or RegJoyas.EOF) Then
'                    MsgBox " No existe la descripción de sus Joyas " & !cCodCta, vbExclamation, " Error del Sistema "
'                    RegJoyas.Close
'                    Set RegJoyas = Nothing
'                    RegCredPrend.Close
'                    Set RegCredPrend = Nothing
'                    Exit Sub
'                Else
'                    Do While Not RegJoyas.EOF
'                        If RegJoyas!ckilataje = "14" Then
'                            v14 = RegJoyas!nPesoOro: v14 = Round(v14, 2)
'                            vPreBas = vPreBas + (v14 * txtPreOro14)
'                            vGramos = vGramos + v14
'                        End If
'                        If RegJoyas!ckilataje = "16" Then
'                            v16 = RegJoyas!nPesoOro: v16 = Round(v16, 2)
'                            vPreBas = vPreBas + (v16 * txtPreOro16)
'                            vGramos = vGramos + v16
'                        End If
'                        If RegJoyas!ckilataje = "18" Then
'                            v18 = RegJoyas!nPesoOro:  v18 = Round(v18, 2)
'                            vPreBas = vPreBas + (v18 * txtPreOro18)
'                            vGramos = vGramos + v18
'                        End If
'                        If RegJoyas!ckilataje = "21" Then
'                            v21 = RegJoyas!nPesoOro:  v21 = Round(v21, 2)
'                            vPreBas = vPreBas + (v21 * txtPreOro21)
'                            vGramos = vGramos + v21
'                        End If
'                        RegJoyas.MoveNext
'                    Loop
'                    RegJoyas.Close
'                    Set RegJoyas = Nothing
'                End If
'                vPreBas = Round(vPreBas, 2)
'                vPreDeu = CalculaDeudaPrendario(!nSaldoCap, !dFecVenc, !nvaltasac, !nTasaIntVenc, _
'                        !nCostCusto, pTasaImpuesto, !cEstado, pTasaPreparacionRemate, txtFecAdjudica)
'                vPreDeu = Round(vPreDeu, 2)
'                'Adjudica al menor entre: Saldo Capital - Precio Oro
'                vPreVen = IIf(!nSaldoCap > vPreBas, vPreBas, !nSaldoCap)
'                vPreVen = Round(vPreVen, 2)
'
'                'Para Ver el Código Antiguo
'                If pVerCodAnt Then vVerCodAnt = IIf(IsNull(!cCodAnt), "", !cCodAnt)
'                If optImpresion(0).Value = True Then
'                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 6, 0) & ImpreFormat(!cCodCta, 13, 1) & Format(!dFecVenc, "dd/mm/yyyy") & ImpreFormat(vNombre, 26, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) & _
'                        ImpreFormat(vGramos, 6) & ImpreFormat(!nvaltasac, 9) & ImpreFormat(!nSaldoCap, 9) & _
'                        ImpreFormat(vPreDeu, 9) & ImpreFormat(vPreBas, 9) & ImpreFormat(vPreVen, 9) & gPrnSaltoLinea
'                    If pVerCodAnt And Len(vVerCodAnt) > 0 Then
'                        vRTFImp = vRTFImp & Space(7) & ImpreFormat(vVerCodAnt, 13, 0) & gPrnSaltoLinea
'                        vLineas = vLineas + 1
'                    End If
'                Else
'                    Print #ArcSal, ImpreFormat(vIndice, 6, 0) & ImpreFormat(!cCodCta, 13, 1) & Format(!dFecVenc, "dd/mm/yyyy") & ImpreFormat(vNombre, 26, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & ImpreFormat(v18, 6) & ImpreFormat(v21, 6) & _
'                        ImpreFormat(vGramos, 6) & ImpreFormat(!nvaltasac, 9) & ImpreFormat(!nSaldoCap, 9) & _
'                        ImpreFormat(vPreDeu, 9) & ImpreFormat(vPreBas, 9) & ImpreFormat(vPreVen, 9)
'                    If pVerCodAnt And Len(vVerCodAnt) > 0 Then
'                        Print #ArcSal, Space(7) & ImpreFormat(vVerCodAnt, 13, 0)
'                        vLineas = vLineas + 1
'                    End If
'                End If
'                vLineas = vLineas + 1
'
'                vSumv14 = vSumv14 + v14
'                vSumv16 = vSumv16 + v16
'                vSumv18 = vSumv18 + v18
'                vSumv21 = vSumv21 + v21
'                vSumGramos = vSumGramos + vGramos
'                vSumTasaci = vSumTasaci + !nvaltasac
'                vSumPresta = vSumPresta + !nSaldoCap
'                vSumPreDeu = vSumPreDeu + vPreDeu
'                vSumPreBas = vSumPreBas + vPreBas
'                vSumPreVen = vSumPreVen + vPreVen
'                vPreBas = 0: vPreDeu = 0: vPreVen = 0
'                v14 = 0: v16 = 0: v18 = 0: v21 = 0: vGramos = 0
'                vIndice = vIndice + 1
'                If vLineas >= 55 Then
'                    vPage = vPage + 1
'                    If optImpresion(0).Value = True Then
'                        vRTFImp = vRTFImp & Chr(12)
'                        Cabecera "PlanPrAd", vPage, False
'                    Else
'                        If vPage Mod 5 = 0 Then
'                            ImpreEnd
'                            ImpreBegin True, pHojaFiMax
'                        Else
'                            ImpreNewPage
'                        End If
'                        vRTFImp = ""
'                        Cabecera "PlanPrAd", vPage, False
'                        Print #ArcSal, ImpreCarEsp(vRTFImp);
'                        vRTFImp = ""
'                    End If
'                    vLineas = 7
'                End If
'                vCont = vCont + 1
'                prgList.Value = vCont
'                .MoveNext
'            Loop
'            If optImpresion(0).Value = True Then
'                vRTFImp = vRTFImp & gPrnSaltoLinea
'                vRTFImp = vRTFImp & Space(24) & "RESUMEN " & ImpreFormat(vSumv14, 31) & ImpreFormat(vSumv16, 6) & ImpreFormat(vSumv18, 6) & ImpreFormat(vSumv21, 6) & _
'                    ImpreFormat(vSumGramos, 6) & ImpreFormat(vSumTasaci, 9) & ImpreFormat(vSumPresta, 9) & _
'                    ImpreFormat(vSumPreDeu, 9) & ImpreFormat(vSumPreBas, 9) & ImpreFormat(vSumPreVen, 9) & gPrnSaltoLinea
'                vRTFImp = vRTFImp & Chr(12)
'            Else
'                Print #ArcSal, ""
'                Print #ArcSal, Space(24) & "RESUMEN " & ImpreFormat(vSumv14, 31) & ImpreFormat(vSumv16, 6) & ImpreFormat(vSumv18, 6) & ImpreFormat(vSumv21, 6) & _
'                    ImpreFormat(vSumGramos, 6) & ImpreFormat(vSumTasaci, 9) & ImpreFormat(vSumPresta, 9) & _
'                    ImpreFormat(vSumPreDeu, 9) & ImpreFormat(vSumPreBas, 9) & ImpreFormat(vSumPreVen, 9)
'                Print #ArcSal, ""
'                ImpreEnd
'            End If
'        End With
'        prgList.Visible = False
'        prgList.Value = 0
'        RegCredPrend.Close
'        Set RegCredPrend = Nothing
'    End If
'    MousePointer = 0
End Sub

'Cabecera de las Impresiones
Private Sub Cabecera(ByVal vOpt As String, ByVal vPagina As Integer, Optional ByVal pPagCorta As Boolean = True)
    Dim vTitulo As String
    Dim vSubTit As String
    Dim vArea As String * 30
    Dim vNroLineas As Integer
    Select Case vOpt
        Case "PlanPrAd"
            vTitulo = "PLANILLA PARA ADJUDICACION N°: " & Format(txtNumAdjudica, "@@@@") & " DEL " & txtFecAdjudica
        Case "ListAdjuMes"
            vTitulo = "LISTADO ADJUDICADOS EN EL MES DE " & UCase(Format(txtFecMes.Text, "mmmm"))
        Case "ListAdju"
            vTitulo = IIf(txtEstado.Text <> "FINALIZADO", "LISTADO CONSOLIDADO DE CONTRATOS ADJUDICADOS", "LISTADO DE ADJUDICACION NRO.: " & List1.Text)
    End Select
    vSubTit = "  "
    vArea = "Crédito Pignoraticio"
    vNroLineas = IIf(pPagCorta = True, 128, 165)
    'Centra Título
    vTitulo = String(Round((vNroLineas - Len(Trim(vTitulo))) / 2), " ") & vTitulo
    'Centra SubTítulo
    vSubTit = String(Round(((vNroLineas - 60) - Len(Trim(vSubTit))) / 2), " ") & vSubTit & String(Round(((vNroLineas - 60) - Len(Trim(vSubTit))) / 2), " ")

    vRTFImp = vRTFImp & gPrnSaltoLinea
    vRTFImp = vRTFImp & Space(1) & ImpreFormat(vNomAge, 25, 0) & Space(vNroLineas - 40) & "Página: " & Format(vPagina, "@@@@") & gPrnSaltoLinea
    vRTFImp = vRTFImp & Space(1) & vTitulo & gPrnSaltoLinea
    vRTFImp = vRTFImp & Space(1) & vArea & vSubTit & Space(7) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm:ss") & gPrnSaltoLinea
    vRTFImp = vRTFImp & String(vNroLineas, "-") & gPrnSaltoLinea
    Select Case vOpt
        Case "PlanPrAd"
            vRTFImp = vRTFImp & Space(1) & "ITEM    CONTRATO     FECHA          NOMBRE CLIENTE          14Kl.    16Kl.    18Kl.    21Kl.  GRAMOS    TASACION      SALDO        DEUDA      PRECIO      VALOR" & gPrnSaltoLinea
            vRTFImp = vRTFImp & Space(1) & "                    VENCIMI.                                                                                          CAPITAL                   ORO      REGISTRO" & gPrnSaltoLinea
        Case "ListAdju", "ListAdjuMes"
            vRTFImp = vRTFImp & Space(1) & "ITEM   CONTRATO     PZ           DESCRIPCION              14Kl.    16Kl.    18Kl.    21Kl.     TOTAL       DEUDA      PRECIO" & gPrnSaltoLinea
            vRTFImp = vRTFImp & Space(1) & "                                                                                                                     ADJUDICA." & gPrnSaltoLinea
    End Select
    vRTFImp = vRTFImp & String(vNroLineas, "-") & gPrnSaltoLinea
End Sub

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnProcesoCadaAgencia As Integer

Set loParam = New COMDColocPig.DCOMColPCalculos
    fnNroRematesAdjudicacion = loParam.dObtieneColocParametro(gConsColPNroRematesParaAdjudic)
    fnTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    fnTasaCustodiaVencida = loParam.dObtieneColocParametro(gConsColPTasaCustodiaVencida)
    fnTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
Set loParam = Nothing
    pPrevioMax = 2000
    pLineasMax = 56
    pHojaFiMax = 66
    'pVerCodAnt = IIf(Left(ReadVarSis("CPR", "cVerCodAnt"), 1) = "S", True, False)
Set loConstSis = New COMDConstSistema.NCOMConstSistema
    lnProcesoCadaAgencia = loConstSis.LeeConstSistema(121)  ' gConstSistPigRemateCadaAg
    If lnProcesoCadaAgencia = 1 Then  ' En cada agencia
        fsProcesoCadaAgencia = gsCodCMAC & gsCodAge
    Else
        fsProcesoCadaAgencia = gsCodCMAC & "00"
    End If
Set loConstSis = Nothing
End Sub

Public Function ExisAdjVenSub(ByVal pCodCta As String) As Boolean
'    Dim tmpReg As New ADODB.Recordset
'    Dim tmpSql As String
'    'Verifica la existencia en DetSubasta
'    tmpSql = " SELECT ccodcta FROM detSubas where ccodcta = '" & pCodCta & "' AND cEstado = 'V'"
'    tmpReg.Open tmpSql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
'    If (tmpReg.BOF Or tmpReg.EOF) Then
'        ExisAdjVenSub = False
'    Else
'        ExisAdjVenSub = True
'    End If
'    tmpReg.Close
'    Set tmpReg = Nothing
End Function

' Permite activar la opción de grabar solo cuando están ingresados los campos
' del precio del oro
Private Sub VeriPreOro()
    If val(txtPreOro14) > 0 And val(txtPreOro16) > 0 And val(txtPreOro18) > 0 _
        And val(txtPreOro21) > 0 Then
        cmdGrabar.Enabled = True
    End If
End Sub
