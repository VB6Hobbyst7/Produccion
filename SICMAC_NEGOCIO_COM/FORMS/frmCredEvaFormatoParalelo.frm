VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredFormEvalFormatoParalelo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creditos - Evaluacion - Formato Paralelo"
   ClientHeight    =   11205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9870
   Icon            =   "frmCredEvaFormatoParalelo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11205
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Evaluacion"
      TabPicture(0)   =   "frmCredEvaFormatoParalelo.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Propuesta del Credito"
      TabPicture(1)   =   "frmCredEvaFormatoParalelo.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame5 
         Caption         =   "Resumen"
         Height          =   2775
         Left            =   5280
         TabIndex        =   29
         Top             =   3360
         Width           =   4095
         Begin VB.TextBox txtMonPropuesto 
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
            Height          =   300
            Left            =   2400
            TabIndex        =   60
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtMonParalelo 
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
            Height          =   300
            Left            =   2400
            TabIndex        =   59
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtResumenIncIngresos 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2400
            TabIndex        =   58
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtIngresos 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2400
            TabIndex        =   57
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtResuMargenBrutoCaja 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2400
            TabIndex        =   56
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label29 
            Caption         =   "Monto Propuesto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   65
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label28 
            Caption         =   "Monto Paralelo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   64
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "Incremento de Ingresos %:"
            Height          =   300
            Left            =   240
            TabIndex        =   63
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label17 
            Caption         =   "Ingresos:"
            Height          =   300
            Left            =   240
            TabIndex        =   62
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "Margen Bruto de Caja:"
            Height          =   300
            Left            =   240
            TabIndex        =   61
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Estimacion Monto"
         Height          =   2775
         Left            =   240
         TabIndex        =   28
         Top             =   3360
         Width           =   4095
         Begin VB.TextBox txtEstMonOtrosIngresos 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2040
            TabIndex        =   48
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtCutCredVigente 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   47
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txtEstMonConsFamiliar 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2040
            TabIndex        =   46
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtEstMonOtrosGasto 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2040
            TabIndex        =   45
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtMagBruto 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   44
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtIncIngreso 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   43
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtEstMonIngreso 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   42
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Cuota Cred. Vigente:"
            Height          =   300
            Left            =   240
            TabIndex        =   55
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   "Otros Ingresos:"
            Height          =   300
            Left            =   240
            TabIndex        =   54
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "Consumo Familiar:"
            Height          =   300
            Left            =   240
            TabIndex        =   53
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "Otros Gastos:"
            Height          =   300
            Left            =   240
            TabIndex        =   52
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "% Margen Bruto:"
            Height          =   300
            Left            =   240
            TabIndex        =   51
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Incremento Ingresos:"
            Height          =   300
            Left            =   240
            TabIndex        =   50
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Ingresos:"
            Height          =   300
            Left            =   240
            TabIndex        =   49
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos"
         Height          =   2295
         Left            =   5280
         TabIndex        =   27
         Top             =   480
         Width           =   4095
         Begin SICMACT.uSpinner spnDatosIncrIngreso 
            Height          =   300
            Left            =   2400
            TabIndex        =   31
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
         End
         Begin VB.Label Label3 
            Caption         =   "Incremento de Ingreso:"
            Height          =   300
            Left            =   240
            TabIndex        =   30
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Propuesta del Credito"
         Height          =   6495
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   9375
         Begin VB.TextBox txtSustentoIncreVenta 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   77
            Top             =   5640
            Width           =   9015
         End
         Begin VB.TextBox txtGarantias 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   75
            Top             =   4680
            Width           =   9015
         End
         Begin VB.TextBox txtFormalidadNegocio 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   73
            Top             =   3720
            Width           =   9015
         End
         Begin VB.TextBox txtCrediticia 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   70
            Top             =   2760
            Width           =   9015
         End
         Begin VB.TextBox txtGiroUbicacion 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   69
            Top             =   1800
            Width           =   9015
         End
         Begin VB.TextBox txtEntornoFamiliar 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   840
            Width           =   9015
         End
         Begin MSMask.MaskEdBox txtFechaVista 
            Height          =   345
            Left            =   7920
            TabIndex        =   67
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label33 
            Caption         =   "Sobre el Destino y el Impacto del Mismo"
            Height          =   300
            Left            =   240
            TabIndex        =   78
            Top             =   5400
            Width           =   4575
         End
         Begin VB.Label Label32 
            Caption         =   "Sobre los Colaterales o Garantias"
            Height          =   300
            Left            =   240
            TabIndex        =   76
            Top             =   4440
            Width           =   3975
         End
         Begin VB.Label Label31 
            Caption         =   "Sobre la Consistencia de la Informacion y la Formalidad del Negocio"
            Height          =   300
            Left            =   240
            TabIndex        =   74
            Top             =   3480
            Width           =   6255
         End
         Begin VB.Label Label30 
            Caption         =   "Sobre la Experiencia Crediticia"
            Height          =   300
            Left            =   240
            TabIndex        =   72
            Top             =   2520
            Width           =   4215
         End
         Begin VB.Label Label27 
            Caption         =   "Sobre el Giro y la Ubicacion del Negocio"
            Height          =   300
            Left            =   240
            TabIndex        =   71
            Top             =   1560
            Width           =   4095
         End
         Begin VB.Label Label2 
            Caption         =   "Sobre el Entorno Familiar del Cliente o Representante"
            Height          =   300
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   4695
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha de Vista:"
            Height          =   300
            Left            =   6720
            TabIndex        =   24
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Datos Credito Vigente"
         Height          =   2295
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   4095
         Begin VB.TextBox txtIngNeto 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2040
            TabIndex        =   36
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtCapPago 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2040
            TabIndex        =   35
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtVentas 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2040
            TabIndex        =   34
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtSaldoActual 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   33
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtMonAprobado 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   32
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Ingreso Neto:"
            Height          =   300
            Left            =   240
            TabIndex        =   41
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Cap. Pago:"
            Height          =   300
            Left            =   240
            TabIndex        =   40
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Ventas:"
            Height          =   300
            Left            =   240
            TabIndex        =   39
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Saldo Actual:"
            Height          =   300
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Monto Aprobado:"
            Height          =   300
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   4080
      TabIndex        =   17
      Top             =   600
      Width           =   5535
      Begin VB.TextBox txtActividad 
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         TabIndex        =   19
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label26 
         Caption         =   "Actividad:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8520
      TabIndex        =   9
      Top             =   10800
      Width           =   1170
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7200
      TabIndex        =   8
      Top             =   10800
      Width           =   1170
   End
   Begin VB.CommandButton cmdInfromeVista 
      Caption         =   "Informe de Vista"
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
      Left            =   120
      TabIndex        =   7
      Top             =   10800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   9375
      Begin VB.TextBox txtFechaExpeCaja 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   79
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboFechaEduSBS 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7320
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtUltimoEduSBS 
         Enabled         =   0   'False
         Height          =   350
         Left            =   7320
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtCampana 
         Enabled         =   0   'False
         Height          =   350
         Left            =   2520
         TabIndex        =   3
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtNCredito 
         Enabled         =   0   'False
         Height          =   350
         Left            =   2520
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtNombCliente 
         Enabled         =   0   'False
         Height          =   350
         Left            =   2520
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
      Begin VB.TextBox txtExpCredito 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   350
         Left            =   7320
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label25 
         Caption         =   "Exposicion con este Credito:"
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label24 
         Caption         =   "Fecha ultimo endeud. SBS:"
         Height          =   255
         Left            =   5160
         TabIndex        =   15
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label23 
         Caption         =   "Ultimo endeudamiento SBS:"
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label22 
         Caption         =   "Campaña:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Numero de Creditos:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label20 
         Caption         =   "Experiencia en la Caja (Desde):"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label19 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3375
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Informacion del Negocio"
      TabPicture(0)   =   "frmCredEvaFormatoParalelo.frx":0342
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ActXCodCta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   495
         Left            =   120
         TabIndex        =   66
         Top             =   600
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         Texto           =   "Credito:"
      End
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7200
      TabIndex        =   68
      Top             =   10800
      Width           =   1170
   End
End
Attribute VB_Name = "frmCredFormEvalFormatoParalelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************************************
'*  Nombre:         frmCredFormEvalFormatoPralelo                                                       *
'*  Descripcion:    Formulario para Evaluacion de Creditos que tiene el tipo de Evaluacion Paralelo     *
'*  Creado:         TI-ERS004-2016                                                                      *
'*  Autor:          JOEP, 25-06-2016                                                                    *
'********************************************************************************************************

Option Explicit
Dim fnTipoCliente As Integer
Dim sCtaCod As String
Dim fnTipoRegMant As Integer
Dim fnTipoPermiso As Integer
Dim fbPermiteGrabar As Boolean
Dim fbBloqueaTodo As Boolean

'Cabecera - Formato Paralelo
Dim fsActividad As String
Dim fsCliente As String
Dim fdPersIng As Date
Dim fnUltimoEduSBS As Double
Dim fnNCred As Integer
Dim fdUltimaEduSBS As Date
Dim fsCampana As String
Dim fnExpCred As Double

'Evaluacion - Formato Paralelo
Dim fnMonAprobado As Double
Dim fnSalActual As Double
Dim fnVentas As Double
Dim fnCapPago As Double
Dim fnIngNeto As Double

Dim fnDatosIncIngreso As Double

Dim fnEstMontoIngresos As Double
Dim fnEstMontoIncIngreso As Double
Dim fnMagBruto As Double
Dim fnOtrGastos As Double
Dim fnConsFamiliar As Double
Dim fnCutCredVigent As Double
Dim fnOtrIngresos As Double

Dim fnMagBrutoCaja As Double
Dim fnIngresos As Double
Dim fnIncIngresos As Double
Dim fnMonParalelo As Double
Dim fnMonPropuesto As Double

'Propuesta del Credito - Formato Paralelo
Dim fdFechaVista As Date
Dim fsSustIncVenta As String

Dim cSPrd As String, cPrd As String
Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
Dim objPista As COMManejador.Pista
Dim nFormato, nPersoneria As Integer
Dim fnMontoIni As Double
Dim lnMin As Double, lnMax As Double
Dim lnMinDol As Double, lnMaxDol As Double
Dim nTC As Double

Public Sub Inicio(ByVal psCtaCod As String, ByVal psTipoRegMant As Integer)

    Call Limpiaformulario
    Call LLenarFormulario
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsDCredito As ADODB.Recordset
    Dim rsDCredEval As ADODB.Recordset
    Dim rsDColCred As ADODB.Recordset
    Dim rsDLLenarEvaluacion As ADODB.Recordset
    
    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
    
    nTC = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia)
    sCtaCod = psCtaCod
    fnTipoRegMant = psTipoRegMant
    
    ActXCodCta.NroCuenta = sCtaCod
      
    '(3: Analista, 2: Coordinador, 1: JefeAgencia)
    fnTipoPermiso = oNCOMFormatosEval.ObtieneTipoPermisoCredEval(gsCodCargo) ' Obtener el tipo de Permiso, Segun Cargo
    
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsDCredito = oDCOMFormatosEval.RecuperarDatosCredEvalFormatoParalelo(sCtaCod) ' Llenar Datos en la Cabecera Informacion de Negocio

    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsDLLenarEvaluacion = oDCOMFormatosEval.RecuperarDatosCredEvalFPEvaluacion(sCtaCod) ' Llenar Datos en Evaluacion

         If fnTipoRegMant = 2 And Mantenimineto(IIf(fnTipoRegMant = 2, False, True)) = False Then
            MsgBox "No Cuenta con Registros", vbInformation, "Aviso"
            Exit Sub
        End If
 

    If Not (rsDCredito.EOF And rsDCredito.BOF) Then
        fsActividad = Trim(rsDCredito!cActiGiro)
        fsCliente = Trim(rsDCredito!cPersNombre)
        fdPersIng = Trim(rsDCredito!dPersIng)
        fnUltimoEduSBS = Trim(rsDCredito!nUltimoEduSBS)
        fnNCred = Trim(rsDCredito!nCreditos)
        fdUltimaEduSBS = Trim(rsDCredito!dFechaUltEnduSBS)
        fsCampana = Trim(rsDCredito!cDescripcion)
        fnExpCred = Trim(rsDCredito!nExpoCred)
               
    End If
    If Not (rsDLLenarEvaluacion.BOF And rsDLLenarEvaluacion.EOF) Then
        fnMonAprobado = Trim(rsDLLenarEvaluacion!nmonto)
        fnSalActual = Trim(rsDLLenarEvaluacion!nSaldo)
        'fnVentas = Trim(rsDLLenarEvaluacion!nImporte)
        'fnCapPago = Trim(rsDLLenarEvaluacion!nCapPago)
        'fnIngNeto = Trim(rsDLLenarEvaluacion!nIngrNeto)
        
    'Estimacion Monto
        'fnEstMontoIngresos
        'fnEstMontoIncIngreso
        'fnMagBruto
        'fnOtrGastos
        'fnConsFamiliar
         fnCutCredVigent = Trim(rsDLLenarEvaluacion!nCuota)
        'fnOtrIngresos
        
    'Resumen
        'fnMagBrutoCaja
        'fnIngresos
        'fnIncIngresos
        'fnMonParalelo
        fnMonPropuesto = Trim(rsDLLenarEvaluacion!nMontoPro)
        
    End If
             
     If Not (rsDCredito.EOF And rsDCredito.BOF) Then
        txtActividad.Text = fsActividad
        txtNombCliente.Text = fsCliente
        txtFechaExpeCaja.Text = fdPersIng
        txtUltimoEduSBS.Text = fnUltimoEduSBS
        txtNCredito.Text = fnNCred
        cboFechaEduSBS.Text = fdUltimaEduSBS
        txtCampana.Text = fsCampana
        txtExpCredito.Text = fnExpCred
    End If
    
    If Not (rsDLLenarEvaluacion.BOF And rsDLLenarEvaluacion.EOF) Then
        txtMonAprobado.Text = Format(fnMonAprobado, "#,##0.00")
        txtSaldoActual.Text = Format(fnSalActual, "#,##0.00")
        'txtVentas.Text = Format(fnVentas, "#,##0.00")
        'txtCapPago.Text = Format(fnCapPago, "#,##0.00")
        'txtIngNeto.Text = Format(fnIngNeto, "#,##0.00")
            
        txtEstMonIngreso.Text = Format(fnVentas, "#,##0.00")
        'txtIncIngreso
        'txtMagBruto
        'txtEstMonOtrosGasto
        'txtEstMonConsFamiliar
        txtCutCredVigente = Format(fnCutCredVigent, "#,##0.00")
        'txtEstMonOtrosIngresos
        
        'txtResuMargenBrutoCaja
        txtIngresos.Text = Format(fnVentas, "#,##0.00")
        'txtResumenIncIngresos
        'txtMonParalelo = fnMonAprobado + A
        txtMonPropuesto.Text = Format(fnMonPropuesto, "#,##0.00")
    End If

'explicar codigo
    'cSPrd = Trim(rsDCredito!cTpoProdCod)
    'cPrd = Mid(cSPrd, 1, 1) & "00"
    'fbPermiteGrabar = False
    'fbBloqueaTodo = False
     
    Set rsDCredEval = oDCOMFormatosEval.RecuperaColocacCredEval(sCtaCod) 'Recuperar Credito Si ha sido Registrado el Form. Eval.
    If fnTipoPermiso = 2 Then
       If rsDCredEval.RecordCount = 0 Then ' Si no hay credito registrado
            MsgBox "El analista no ha registrado la Evaluacion respectiva", vbExclamation, "Aviso"
            fbPermiteGrabar = False
        Else
            fbPermiteGrabar = True
         End If
    End If
    Set rsDCredito = Nothing
    Set rsDCredEval = Nothing
    
    
    
    Set rsDColCred = oDCOMFormatosEval.RecuperaColocacCred(sCtaCod) ' PARA VERFICAR SI FUE VERIFICADO
'    If rsDColCred!nVerifCredEval = 1 Then
'        MsgBox "Ud. no puede editar la evaluación, ya se realizó la verificacion del credito", vbExclamation, "Aviso"
'        fbBloqueaTodo = True
'    End If
    nFormato = oDCOMFormatosEval.AsignarFormato(cPrd, cSPrd, fnMontoIni)
    'lnMinDol = lnMin / nTC 'Convertimos al tipo de cambio
    'lnMaxDol = lnMax / nTC
    
    Set oDCOMFormatosEval = Nothing
    Set oTipoCam = Nothing
'    If CargaDatos Then
'        If CargaControles(fnTipoPermiso, fbPermiteGrabar, fbBloqueaTodo) Then
'            If fnTipoRegMant = 1 Then 'Para el Evento: "Registrar"
'                If Not rsCredEval.EOF Then
'                    'Call Mantenimiento
'                    fnTipoRegMant = 2
'                Else
'                    Call Registro
'                    fnTipoRegMant = 1
'                End If
'            Else ' Para el Evento. "Mantenimiento"
'                If rsCredEval.EOF Then
'                    Call Registro
'                    fnTipoRegMant = 1
'                Else
'                    'Call Mantenimiento
'                    fnTipoRegMant = 2
'                End If
'            End If
'        Else
'            Unload Me
'            Exit Sub
'        End If
'    Else
'        If CargaControles(1, False) Then
'        End If
'    End If
    If txtFechaVista.Text <> "__/__/____" And fnTipoRegMant = 1 Then
        MsgBox "Ya cuenta con una evaluación"
    Else
        Me.Show 1
    End If
End Sub

'Actualizar Datos
Private Sub cmdActualizar_Click()
    Dim oCredFormEval As COMNCredito.NCOMFormatosEval
    
    Dim ActualizarFormatoParalelo As Boolean
   
    
    If ValidarDatosFormatoParalelo Then
       
        Set oCredFormEval = New COMNCredito.NCOMFormatosEval
               
        ActualizarFormatoParalelo = oCredFormEval.ActualizarfrmCredFormEvalFormatoParalelo(sCtaCod, CDate(txtFechaExpeCaja.Text), _
                                                        txtMonAprobado.Text, txtSaldoActual.Text, txtVentas.Text, txtCapPago.Text, txtIngNeto.Text, _
                                                        spnDatosIncrIngreso.valor, _
                                                        txtEstMonIngreso.Text, txtIncIngreso.Text, txtMagBruto.Text, txtEstMonOtrosGasto.Text, txtEstMonConsFamiliar.Text, txtCutCredVigente.Text, txtEstMonOtrosIngresos.Text, _
                                                        txtResuMargenBrutoCaja.Text, txtIngresos.Text, txtResumenIncIngresos.Text, txtMonParalelo.Text, txtMonPropuesto.Text, _
                                                        CDate(txtFechaVista.Text), Trim(txtSustentoIncreVenta.Text), txtEntornoFamiliar.Text, txtGiroUbicacion.Text, txtCrediticia.Text, txtFormalidadNegocio.Text, txtGarantias.Text)
        
        
        
        
        
        
        If ActualizarFormatoParalelo Then
            MsgBox "Los Datos se Actualizaron Correctamente"
            Else
                MsgBox "Hubo error al grabar la informacion", vbError, "Error"
        End If
        cmdInfromeVista.Enabled = True
        Controles
        cmdActualizar.Enabled = False
    End If
End Sub

'Salir
Private Sub cmdCancelar_Click()
Unload Me
End Sub
'Guardar Datos
Private Sub cmdGuardar_Click()
    Dim oCredFormEval As COMNCredito.NCOMFormatosEval
    Dim GrabarFormatoParalelo As Boolean
    Dim rsEvaluacion As ADODB.Recordset
    
    If ValidarDatosFormatoParalelo Then
        Set rsEvaluacion = LenarRecordset_Evaluacion
        Set oCredFormEval = New COMNCredito.NCOMFormatosEval
        'Set oCred = New COMNCredito.NCOMCredito
        
        If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        GrabarFormatoParalelo = oCredFormEval.GrabarfrmCredFormEvalFormatoParalelo(sCtaCod, 9, Trim(txtActividad.Text), CDate(txtFechaExpeCaja.Text), _
                                                                            txtUltimoEduSBS.Text, txtNCredito.Text, CDate(cboFechaEduSBS.Text), _
                                                                            Trim(txtCampana.Text), txtExpCredito.Text, _
                                                                            rsEvaluacion, txtVentas, txtCapPago, txtIngNeto, _
                                                                            spnDatosIncrIngreso.valor, txtEstMonIngreso.Text, txtIncIngreso.Text, txtMagBruto.Text, txtEstMonOtrosGasto.Text, txtEstMonConsFamiliar.Text, txtCutCredVigente.Text, _
                                                                            txtEstMonOtrosIngresos.Text, txtResuMargenBrutoCaja.Text, txtIngresos.Text, txtResumenIncIngresos, txtMonParalelo, _
                                                                            CDate(txtFechaVista.Text), Trim(txtSustentoIncreVenta.Text), txtEntornoFamiliar.Text, txtGiroUbicacion.Text, txtCrediticia.Text, txtFormalidadNegocio.Text, txtGarantias.Text)
                        
        If GrabarFormatoParalelo Then
            MsgBox "Los Datos se Grabaron Correctamente"
            Else
                MsgBox "Hubo error al grabar la informacion", vbError, "Error"
        End If
        cmdInfromeVista.Enabled = True
        Controles
    End If
    
End Sub

Public Sub Controles()

txtFechaExpeCaja.Enabled = False
txtVentas.Enabled = False
txtCapPago.Enabled = False
txtIngNeto.Enabled = False
spnDatosIncrIngreso.Enabled = False
txtEstMonOtrosGasto.Enabled = False
txtEstMonConsFamiliar.Enabled = False
txtEstMonOtrosIngresos.Enabled = False
txtResuMargenBrutoCaja.Enabled = False
txtFechaVista.Enabled = False
txtSustentoIncreVenta.Enabled = False

cmdGuardar.Enabled = False
cmdActualizar = False

End Sub

Public Function LenarRecordset_Evaluacion() As ADODB.Recordset

Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
Dim rsEvaluacion As ADODB.Recordset

Set rsEvaluacion = New ADODB.Recordset
Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval

Set rsEvaluacion = oDCOMFormatosEval.RecuperarDatosCredEvalFPEvaluacion(sCtaCod) ' llenar mi formato evaluacion
    
fnMonAprobado = Trim(rsEvaluacion!nmonto)
fnSalActual = Trim(rsEvaluacion!nSaldo)
'fnVentas = Trim(rsEvaluacion!nImporte)
'fnCapPago = Trim(rsEvaluacion!nCapPago)
'fnIngNeto = Trim(rsEvaluacion!nIngrNeto)

'fnDatosIncIngreso = Trim(rsEvaluacion!nIncreIngresoDatos)

'fnEstMontoIngresos = Trim(rsEvaluacion!nIngreEstMontos)
'fnEstMontoIncIngreso = Trim(rsEvaluacion!nIncreIngresoEstiMontos)
'fnMagBruto = Trim(rsEvaluacion!nMargenbruto)
'fnOtrGastos = Trim(rsEvaluacion!nOtrsoGastos)
'fnConsFamiliar = Trim(rsEvaluacion!nConsuFamili)
'fnCutCredVigent = Trim(rsEvaluacion!nCuota)
'fnOtrIngresos = Trim(rsEvaluacion!nOtrosIng)

'fnMagBrutoCaja = Trim(rsEvaluacion!nMargenBrutoCaja)
'fnIngresos = Trim(rsEvaluacion!nIngreResumen)
'fnIncIngresos = Trim(rsEvaluacion!nIncreIngreResumen)
'fnMonParalelo = Trim(rsEvaluacion!nMontoParalelo)
fnMonPropuesto = Trim(rsEvaluacion!nMontoPro)

Set LenarRecordset_Evaluacion = rsEvaluacion
End Function

Public Function Mantenimineto(ByVal pbMantenimiento As Boolean) As Boolean

Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
Dim rsMantenimientoFormatoParalelo As ADODB.Recordset
Mantenimineto = False

Set rsMantenimientoFormatoParalelo = New ADODB.Recordset
 
Set rsMantenimientoFormatoParalelo = oDCOMFormatosEval.RecuperarDatosTotalFormatoParalelo(sCtaCod)

    If Not (rsMantenimientoFormatoParalelo.BOF And rsMantenimientoFormatoParalelo.EOF) Then
        txtActividad.Text = rsMantenimientoFormatoParalelo!cActividad
        txtNombCliente.Text = rsMantenimientoFormatoParalelo!cPersNombre
        txtFechaExpeCaja.Text = rsMantenimientoFormatoParalelo!dFechaExpeCaja
        txtUltimoEduSBS.Text = rsMantenimientoFormatoParalelo!nUltEndeSBS
        txtNCredito.Text = rsMantenimientoFormatoParalelo!nNCreditos
        cboFechaEduSBS.Text = rsMantenimientoFormatoParalelo!dUltEndeuSBS
        txtCampana.Text = rsMantenimientoFormatoParalelo!cCampaña
        txtExpCredito.Text = rsMantenimientoFormatoParalelo!nExposiCred
        
        txtMonAprobado.Text = Format(rsMantenimientoFormatoParalelo!nMontoApro, "#,##0.00")
        txtSaldoActual.Text = Format(rsMantenimientoFormatoParalelo!nSaldoActual, "#,##0.00")
        txtVentas.Text = Format(rsMantenimientoFormatoParalelo!nVentas, "#,##0.00")
        txtCapPago.Text = Format(rsMantenimientoFormatoParalelo!nCapPago, "#,##0.00")
        txtIngNeto.Text = Format(rsMantenimientoFormatoParalelo!nIngresoNeto, "#,##0.00")
        
        spnDatosIncrIngreso.valor = rsMantenimientoFormatoParalelo!nIncreIngresoDatos
        
        txtEstMonIngreso.Text = Format(rsMantenimientoFormatoParalelo!nIngreEstMontos, "#,##0.00")
        txtIncIngreso.Text = Format(rsMantenimientoFormatoParalelo!nIncreIngresoEstiMontos, "#,##0.00")
        txtMagBruto.Text = Format(rsMantenimientoFormatoParalelo!nMargenBruto, "#,##0.00")
        txtEstMonOtrosGasto.Text = Format(rsMantenimientoFormatoParalelo!nOtrsoGastos, "#,##0.00")
        txtEstMonConsFamiliar.Text = Format(rsMantenimientoFormatoParalelo!nConsuFamili, "#,##0.00")
        txtCutCredVigente.Text = Format(rsMantenimientoFormatoParalelo!nCuotaCredVig, "#,##0.00")
        txtEstMonOtrosIngresos.Text = Format(rsMantenimientoFormatoParalelo!nOtrosIng, "#,##0.00")
        
        txtResuMargenBrutoCaja.Text = Format(rsMantenimientoFormatoParalelo!nMargenBrutoCaja, "#,##0.00")
        txtIngresos.Text = Format(rsMantenimientoFormatoParalelo!nIngreResumen, "#,##0.00")
        txtResumenIncIngresos.Text = Format(rsMantenimientoFormatoParalelo!nIncreIngreResumen, "#,##0.00")
        txtMonParalelo.Text = Format(rsMantenimientoFormatoParalelo!nMontoParalelo, "#,##0.00")
        txtMonPropuesto.Text = Format(rsMantenimientoFormatoParalelo!nMontoPropuesto, "#,##0.00")
        
        txtFechaVista.Text = rsMantenimientoFormatoParalelo!dFecVisita
        txtSustentoIncreVenta.Text = rsMantenimientoFormatoParalelo!cSustenIncreVta
        txtEntornoFamiliar.Text = rsMantenimientoFormatoParalelo!cEntornoFami
        txtGiroUbicacion.Text = rsMantenimientoFormatoParalelo!cExpelaboral
        txtCrediticia.Text = rsMantenimientoFormatoParalelo!cExpeCrediticia
        txtFormalidadNegocio.Text = rsMantenimientoFormatoParalelo!cColateGarantia
        txtGarantias.Text = rsMantenimientoFormatoParalelo!cDestino
        
        Mantenimineto = True
    End If
    cmdGuardar.Visible = pbMantenimiento
    cmdActualizar.Visible = Not pbMantenimiento
End Function

'validar Datos

Public Function ValidarDatosFormatoParalelo() As Boolean

ValidarDatosFormatoParalelo = True

If txtFechaExpeCaja.Text = "__/__/____" Then
    MsgBox "Ingrese Fecha de Experiencia en la Caja ", vbInformation, "Aviso"
    SSTab1.Tab = 0
    ValidarDatosFormatoParalelo = False
    Exit Function
End If

If val(spnDatosIncrIngreso.valor) = 0 Then
    MsgBox "Seleccione Incremento de Ingreso ", vbInformation, "Aviso"
    SSTab1.Tab = 0
    ValidarDatosFormatoParalelo = False
    Exit Function
End If

If Trim(txtEstMonOtrosGasto.Text) = "" Then
    MsgBox "Falta Ingresar Otros Gastos", vbInformation, "Aviso"
    SSTab1.Tab = 0
    ValidarDatosFormatoParalelo = False
    Exit Function
End If
If Trim(txtEstMonConsFamiliar.Text) = "" Then
    MsgBox "Falta Ingresar Consumo Familiar", vbInformation, "Aviso"
    SSTab1.Tab = 0
    ValidarDatosFormatoParalelo = False
    Exit Function
End If

If Trim(txtEstMonOtrosIngresos.Text) = "" Then
    MsgBox "Falta Ingresar Otros Ingresos", vbInformation, "Aviso"
    SSTab1.Tab = 0
    ValidarDatosFormatoParalelo = False
    Exit Function
End If

If Trim(txtResuMargenBrutoCaja.Text) = "" Then
    MsgBox "Falta Ingresar Margen Bruto", vbInformation, "Aviso"
    SSTab1.Tab = 0
    ValidarDatosFormatoParalelo = False
    Exit Function
End If

If txtFechaVista.Text = "__/__/____" Then
    MsgBox "Ingresar Fecha de Vista", vbInformation, "Aviso"
    SSTab1.Tab = 1
    ValidarDatosFormatoParalelo = False
    Exit Function
End If

If Trim(txtSustentoIncreVenta.Text) = "" Then
    MsgBox "Falta Ingresar un Sustento de Incremento de Venta", vbInformation, "Aviso"
    SSTab1.Tab = 1
    ValidarDatosFormatoParalelo = False
    Exit Function
End If

End Function

Private Sub cmdInfromeVista_Click()

Dim rsImformeVisitaFormatoParalelo As ADODB.Recordset

Dim oDoc  As cPDF
Dim psCtaCod As String
Set oDoc = New cPDF

Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
Set rsImformeVisitaFormatoParalelo = New ADODB.Recordset

Set rsImformeVisitaFormatoParalelo = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormatoParalelo(sCtaCod)

'Creación del Archivo
oDoc.Author = gsCodUser
oDoc.Creator = "SICMACT - Negocio"
oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
oDoc.Subject = "Pagaré de Crédito Nº " & sCtaCod
oDoc.Title = "Pagaré de Crédito Nº " & sCtaCod

If Not oDoc.PDFCreate(App.Path & "\Spooler\FORMATOPARALELO_" & sCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
    Exit Sub
End If

'Contenido

oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding

oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
    
'Tamaño de hoja A4
oDoc.NewPage A4_Vertical
'35
oDoc.WImage 40, 60, 35, 105, "Logo"
oDoc.WTextBox 40, 60, 35, 390, rsImformeVisitaFormatoParalelo!cAgeDescripcion, "F2", 10, hLeft

oDoc.WTextBox 40, 60, 35, 390, "FECHA", "F2", 10, hRight
oDoc.WTextBox 40, 60, 35, 450, Format(gdFecSis, "dd/mm/yyyy"), "F2", 10, hRight
oDoc.WTextBox 40, 60, 35, 490, Format(Time, "hh:mm:ss"), "F2", 10, hRight
Dim B As Integer
B = 29
oDoc.WTextBox 90 - B, 60, 15, 160, "Cliente", "F2", 10, hLeft
oDoc.WTextBox 90 - B, 60, 15, 80, ":", "F2", 10, hRight
oDoc.WTextBox 90 - B, 150, 15, 500, rsImformeVisitaFormatoParalelo!cPersNombre, "F1", 10, hjustify

oDoc.WTextBox 90 - B, 350, 15, 160, "Analista", "F2", 10, hjustify
oDoc.WTextBox 90 - B, 390, 15, 80, ":", "F2", 10, hjustify
oDoc.WTextBox 90 - B, 402, 15, 500, rsImformeVisitaFormatoParalelo!cUser, "F1", 10, hjustify

oDoc.WTextBox 100 - B, 60, 15, 160, "Usuario", "F2", 10, hLeft
oDoc.WTextBox 100 - B, 60, 15, 80, ":", "F2", 10, hRight
oDoc.WTextBox 100 - B, 150, 15, 118, gsCodUser, "F1", 10, hjustify

oDoc.WTextBox 100 - B, 350, 15, 160, "Producto", "F2", 10, hjustify
oDoc.WTextBox 100 - B, 390, 15, 80, ":", "F2", 10, hjustify
oDoc.WTextBox 100 - B, 402, 15, 118, rsImformeVisitaFormatoParalelo!cConsDescripcion, "F1", 10, hjustify

oDoc.WTextBox 110 - B, 60, 15, 160, "Credito", "F2", 10, hLeft
oDoc.WTextBox 110 - B, 60, 15, 80, ":", "F2", 10, hRight
oDoc.WTextBox 110 - B, 150, 15, 500, rsImformeVisitaFormatoParalelo!cCtaCod, "F1", 10, hjustify

oDoc.WTextBox 120 - B, 60, 15, 160, "Cod. Cliente", "F2", 10, hLeft
oDoc.WTextBox 120 - B, 60, 15, 80, ":", "F2", 10, hRight
oDoc.WTextBox 120 - B, 150, 15, 500, rsImformeVisitaFormatoParalelo!cPersCod, "F1", 10, hjustify

oDoc.WTextBox 120 - B, 270, 15, 160, "Doc. Natural", "F2", 10, hjustify
oDoc.WTextBox 120 - B, 328, 15, 80, ":", "F2", 10, hjustify
oDoc.WTextBox 120 - B, 335, 15, 500, rsImformeVisitaFormatoParalelo!cPersDni, "F1", 10, hjustify

oDoc.WTextBox 120 - B, 400, 15, 160, "Doc. Juridico", "F2", 10, hjustify
oDoc.WTextBox 120 - B, 460, 15, 80, ":", "F2", 10, hjustify
oDoc.WTextBox 120 - B, 470, 15, 500, rsImformeVisitaFormatoParalelo!cPersRuc, "F1", 10, hjustify

Dim A As Integer
A = 50
            'bajar izq  ar  der
oDoc.WTextBox 110, 100, 15, 400, "INFORME DE VISITA AL CLIENTE", "F2", 12, hCenter

'cuadro de Fecha de visita
oDoc.WTextBox 130, 50, 80, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
'135
oDoc.WTextBox 185 - A, 55, 15, 160, "Fecha de Visita :", "F1", 10, hLeft
oDoc.WTextBox 185 - A, 190, 15, 500, Format(rsImformeVisitaFormatoParalelo!dFechaVisita, "dd/mm/yyyy"), "F1", 10, hjustify

oDoc.WTextBox 200 - A, 55, 15, 160, "Fecha de ultima visita :", "F1", 10, hLeft
oDoc.WTextBox 215 - A, 55, 15, 160, "Persona(s) Entrevistada(s) :", "F1", 10, hLeft

oDoc.WTextBox 230 - A, 55, 15, 160, "Sr.(a) :", "F1", 10, hLeft
oDoc.WTextBox 230 - A, 300, 15, 160, "Cargo/Parentesco :", "F1", 10, hjustify
oDoc.WTextBox 245 - A, 55, 15, 160, "Sr.(a) :", "F1", 10, hLeft
oDoc.WTextBox 245 - A, 300, 15, 160, "Cargo/Parentesco :", "F1", 10, hjustify
'oDoc.WTextBox 260, 55, 15, 160, "Sr.(a) :", "F1", 10, hLeft
'oDoc.WTextBox 260, 350, 15, 160, "Cargo :", "F1", 10, hjustify

'cuadro de Tipo de Visita
oDoc.WTextBox 260 - A, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack

oDoc.WTextBox 262 - A, 55, 15, 500, "Tipo de Visita :", "F2", 10, hLeft
'
'cuadro de Tipo de Visita: Contenido
oDoc.WTextBox 275 - A, 50, 40, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack

oDoc.WTextBox 280 - A, 55, 15, 10, "( )", "F1", 10, hjustify
oDoc.WTextBox 280 - A, 75, 15, 500, "1° Evaluacion (Cliente Nuevo)", "F1", 10, hjustify
oDoc.WTextBox 280 - A, 250, 15, 500, "( )", "F1", 10, hjustify
oDoc.WTextBox 280 - A, 270, 15, 500, "Paralelo", "F1", 10, hjustify
oDoc.WTextBox 280 - A, 400, 15, 700, "( )", "F1", 10, hjustify
oDoc.WTextBox 280 - A, 420, 15, 800, "Inspeccion de Garantias", "F1", 10, hjustify
oDoc.WTextBox 295 - A, 55, 15, 900, "( )", "F1", 10, hjustify
oDoc.WTextBox 295 - A, 75, 15, 110, "Represtamo", "F1", 10, hjustify
oDoc.WTextBox 295 - A, 250, 15, 120, "( )", "F1", 10, hjustify
oDoc.WTextBox 295 - A, 270, 15, 130, "Ampliacion", "F1", 10, hjustify


'cuadro de Sobre el Entorno Familiar del Cliente o Representante
oDoc.WTextBox 315 - A, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack

oDoc.WTextBox 317 - A, 55, 15, 500, "Sobre el Entorno Familiar del Cliente o Representante:", "F2", 10, hLeft

'cuadro de Sobre el Entorno Familiar del Cliente o Representante : CONTENIDO
oDoc.WTextBox 330 - A, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
oDoc.WTextBox 335 - A, 55, 10, 500, rsImformeVisitaFormatoParalelo!cEntornoFami, "F1", 10, hjustify

'cuadro de Sobre el giro y la Ubicacion del Negocio
oDoc.WTextBox 380 - A, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack

oDoc.WTextBox 382 - A, 55, 15, 500, "Sobre el Giro y la Ubicacion del Negocio:", "F2", 10, hLeft

'cuadro de Sobre el giro y la Ubicacion del Negocio : CONTENIDO
oDoc.WTextBox 395 - A, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
oDoc.WTextBox 397 - A, 55, 10, 500, rsImformeVisitaFormatoParalelo!cExpelaboral, "F1", 10, hjustify

'cuadro de Sobre la Experiencia Crediticia
oDoc.WTextBox 445 - A, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
oDoc.WTextBox 447 - A, 55, 15, 500, "Sobre la Experiencia Crediticia:", "F2", 10, hLeft

'cuadro de Sobre la Experiencia Crediticia : CONTENIDO
oDoc.WTextBox 460 - A, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
oDoc.WTextBox 462 - A, 55, 10, 500, rsImformeVisitaFormatoParalelo!cExpeCrediticia, "F1", 10, hjustify

'cuadro de Sobre la consistencia de la informacion y la formalidad del negocio
oDoc.WTextBox 510 - A, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
oDoc.WTextBox 512 - A, 55, 15, 500, "Sobre la Consistencia de la Informacion y la Formalidad del Negocio:", "F2", 10, hLeft

'cuadro de Sobre la consistencia de la informacion y la formalidad del negocio : CONTENIDO
oDoc.WTextBox 525 - A, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
oDoc.WTextBox 527 - A, 55, 10, 500, rsImformeVisitaFormatoParalelo!cColateGarantia, "F1", 10, hjustify

'cuadro de Sobre la Colaterales o Garantias
oDoc.WTextBox 575 - A, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
oDoc.WTextBox 577 - A, 55, 15, 500, "Sobre los Colaterales o Garantias:", "F2", 10, hLeft

'cuadro de Sobre la Colaterales o Garantias : CONTENIDO
oDoc.WTextBox 590 - A, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
oDoc.WTextBox 592 - A, 55, 10, 500, rsImformeVisitaFormatoParalelo!cDestino, "F1", 10, hjustify

'cuadro de Sustento de Venta
oDoc.WTextBox 640 - A, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
oDoc.WTextBox 642 - A, 55, 15, 500, "Sustento de venta:", "F2", 10, hLeft

'cuadro de Sustento de Venta : CONTENIDO
oDoc.WTextBox 655 - A, 50, 50, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
oDoc.WTextBox 657 - A, 55, 10, 500, rsImformeVisitaFormatoParalelo!cSustenIncreVta, "F1", 10, hjustify

'cuadro de VERIFICACION DE INMUEBLE
oDoc.WTextBox 705 - A, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
oDoc.WTextBox 707 - A, 55, 15, 500, "Verificacion de Inmueble :", "F2", 10, hLeft

'cuadro de VERIFICACION DE INMUEBLE:coNTENIDO
oDoc.WTextBox 720 - A, 50, 95, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack

oDoc.WTextBox 722 - A, 55, 15, 500, "Direccion :", "F1", 10, hLeft
oDoc.WTextBox 735 - A, 55, 15, 500, "Referencia de Ubicacion :", "F1", 10, hLeft
oDoc.WTextBox 747 - A, 55, 15, 500, "Zona :", "F1", 10, hLeft
oDoc.WTextBox 755 - A, 200, 15, 500, "( )", "F1", 10, hjustify
oDoc.WTextBox 755 - A, 220, 50, 500, "Urbana", "F1", 10, hjustify
oDoc.WTextBox 755 - A, 280, 60, 500, "( )", "F1", 10, hjustify
oDoc.WTextBox 755 - A, 300, 70, 500, "Rural", "F1", 10, hjustify
oDoc.WTextBox 767 - A, 55, 15, 500, "Tipo de Construccion :", "F1", 10, hLeft
oDoc.WTextBox 780 - A, 100, 15, 500, "( )", "F1", 10, hjustify
oDoc.WTextBox 780 - A, 120, 15, 500, "Material Noble", "F1", 10, hjustify
oDoc.WTextBox 780 - A, 200, 15, 500, "( )", "F1", 10, hjustify
oDoc.WTextBox 780 - A, 220, 15, 500, "Madera", "F1", 10, hjustify
oDoc.WTextBox 780 - A, 280, 15, 500, "( )", "F1", 10, hjustify
oDoc.WTextBox 780 - A, 300, 15, 500, "Otros", "F1", 10, hjustify
oDoc.WTextBox 795 - A, 55, 15, 500, "Estado de la Vivienda :", "F1", 10, hLeft

'cuadro de VISTO BUENO
oDoc.WTextBox 815 - A, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack

oDoc.WTextBox 817 - A, 55, 15, 500, "Analista de Creditos :", "F2", 10, hjustify
oDoc.WTextBox 817 - A, 320, 15, 500, "Jefe de Grupo :", "F2", 10, hjustify

'cuadro de VISTO BUENO:Contenido
oDoc.WTextBox 780, 50, 40, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack

oDoc.PDFClose
oDoc.Show
End Sub

'Pasar el Valor del TextBox a otro Textbox
Private Sub spnDatosIncrIngreso_LostFocus()
    txtEstMonOtrosGasto.SetFocus
    txtResumenIncIngresos.Text = Format(CDbl(spnDatosIncrIngreso.valor), "#,#0.00")
End Sub
'FIN Pasar el Valor del TextBox a otro Textbox

'Al momento de apretar Enter se Va a otro TEXTBOX
Private Sub spnDatosIncrIngreso_KeyPress(KeyAscii As Integer)
 If val(spnDatosIncrIngreso.valor) > 100 Then
        MsgBox "El valor no Puede ser Mayor de 100", vbInformation, "Aviso"
        spnDatosIncrIngreso.valor = 0
    ElseIf KeyAscii = 13 Then
        EnfocaControl txtEstMonOtrosGasto
 End If
  fEnfoque txtEstMonOtrosGasto

    Call CalculoTotal(3)
End Sub

Private Sub txtCapPago_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCapPago, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
'para direccionar a otro campo con apretar el boton "ENTER"
        EnfocaControl txtIngNeto
    End If
        fEnfoque txtIngNeto
End Sub

Private Sub txtCrediticia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
'para direccionar a otro campo con apretar el boton "ENTER"
        EnfocaControl txtFormalidadNegocio
    End If
End Sub

Private Sub txtEntornoFamiliar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtGiroUbicacion
    End If
End Sub

Private Sub txtEstMonOtrosGasto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtEstMonOtrosGasto, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
'para direccionar a otro campo con apretar el boton "ENTER"
        EnfocaControl txtEstMonConsFamiliar
    End If
        fEnfoque txtEstMonConsFamiliar
End Sub
Private Sub txtEstMonConsFamiliar_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEstMonConsFamiliar, KeyAscii, 10, , True)
        If KeyAscii = 13 Then
         SendKeys "{Tab}", True
'para direccionar a otro campo con apretar el boton "ENTER"
            EnfocaControl txtEstMonOtrosIngresos
        End If
        fEnfoque txtEstMonOtrosIngresos
End Sub
Private Sub txtEstMonOtrosIngresos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEstMonOtrosIngresos, KeyAscii, 10, , True)
        If KeyAscii = 13 Then
        SendKeys "{Tab}", True
            EnfocaControl txtResuMargenBrutoCaja
        End If
        fEnfoque txtResuMargenBrutoCaja
End Sub

Private Sub txtVentas_GotFocus()
''Me.txtFechaExpeCaja.SelStart = 0
Me.txtVentas.SelLength = Len(txtVentas.Text)
End Sub

Private Sub txtFormalidadNegocio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
            EnfocaControl txtGarantias
        End If
End Sub

Private Sub txtGarantias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
            EnfocaControl txtSustentoIncreVenta
    End If
End Sub

Private Sub txtGiroUbicacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
            EnfocaControl txtCrediticia
        End If
End Sub

Private Sub txtIngNeto_KeyPress(KeyAscii As Integer)
 KeyAscii = NumerosDecimales(txtIngNeto, KeyAscii, 10, , True)
        If KeyAscii = 13 Then
            spnDatosIncrIngreso.SetFocus
        End If
        
End Sub

Private Sub txtVentas_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtVentas, KeyAscii, 10, , True)
        If KeyAscii = 13 Then
        SendKeys "{Tab}", True
            EnfocaControl txtCapPago
        End If
            fEnfoque txtCapPago
End Sub

Private Sub txtFechaVista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtEntornoFamiliar
    End If
End Sub

Private Sub txtResuMargenBrutoCaja_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtResuMargenBrutoCaja, KeyAscii, 10, , True)
        If KeyAscii = 13 Then
        SendKeys "{Tab}", True
            SSTab1.Tab = 0
            'EnfocaControl txtFechaVista
        End If
End Sub
'fin 'Al momento de apretar Enter se Va a otro TEXTBOX ***********************

'Asignar Decimales a los TEXTBOX
Private Sub txtEstMonConsFamiliar_LostFocus()
    
    If Len(Trim(txtEstMonConsFamiliar.Text)) = "" Then
        txtEstMonConsFamiliar.Text = "0.00"
    End If
        txtEstMonConsFamiliar.Text = Format(txtEstMonConsFamiliar.Text, "#,##0.00")
        
        Call CalculoTotal(2)
End Sub
Private Sub txtEstMonOtrosGasto_LostFocus()
    If Len(Trim(txtEstMonOtrosGasto.Text)) = "" Then
        txtEstMonOtrosGasto.Text = "0.00"
    End If
        txtEstMonOtrosGasto.Text = Format(txtEstMonOtrosGasto.Text, "#,##0.00")
        
        Call CalculoTotal(2)
End Sub
Private Sub txtEstMonOtrosIngresos_LostFocus()
    If Len(Trim(txtEstMonOtrosIngresos.Text)) = "" Then
        txtEstMonOtrosIngresos.Text = "0.00"
    End If
        txtEstMonOtrosIngresos.Text = Format(txtEstMonOtrosIngresos.Text, "#,##0.00")
        
        Call CalculoTotal(2)
End Sub
Private Sub txtResuMargenBrutoCaja_LostFocus()
    If Trim(txtResuMargenBrutoCaja.Text) = "" Then
        txtResuMargenBrutoCaja.Text = "0.00"
        Else
        txtResuMargenBrutoCaja.Text = Format(txtResuMargenBrutoCaja.Text, "###," & String(15, "#") & "#,##0.00")
    End If
    
        Call CalculoTotal(1)
    
        Call CalculoTotal(2)
        
        Call CalculoTotal(4)
    
End Sub
' fin Asignar Decimales a los TEXTBOX

Public Sub Form_Load()
cmdInfromeVista.Enabled = False
cmdActualizar.Visible = False
CentraForm Me
End Sub

Private Sub Limpiaformulario()
  
    txtMonAprobado.Text = ""
    txtSaldoActual.Text = ""
    txtVentas.Text = ""
    txtCapPago.Text = ""
    txtIngNeto.Text = ""
        
    spnDatosIncrIngreso.valor = 0
        
    txtEstMonIngreso.Text = ""
    txtIncIngreso.Text = ""
    txtMagBruto.Text = ""
    txtEstMonOtrosGasto.Text = ""
    txtEstMonConsFamiliar.Text = ""
    txtCutCredVigente.Text = ""
    txtEstMonOtrosIngresos.Text = ""
        
    txtResuMargenBrutoCaja.Text = ""
    txtIngresos.Text = ""
    txtResumenIncIngresos.Text = ""
    txtMonParalelo.Text = ""
    txtMonPropuesto.Text = ""
        
    txtFechaVista.Text = "__/__/____"
    txtSustentoIncreVenta.Text = ""
End Sub

Private Sub LLenarFormulario()
           
    txtVentas.Text = "0.00"
    txtCapPago.Text = "0.00"
    txtIngNeto.Text = "0.00"
    
    spnDatosIncrIngreso.valor = "0.00"
        
    txtMagBruto.Text = "0.00"
    txtIncIngreso.Text = "0.00"
    txtEstMonOtrosGasto.Text = "0.00"
    txtEstMonConsFamiliar.Text = "0.00"
    txtEstMonOtrosIngresos.Text = "0.00"
        
    txtResuMargenBrutoCaja.Text = "0.00"
    txtResumenIncIngresos.Text = "0.00"
    txtMonParalelo.Text = "0.00"
End Sub

Private Sub CalculoTotal(ByVal pnTipo As Integer)
On Error GoTo ErrorCalculo

    Select Case pnTipo
    Case 1:
            txtMagBruto.Text = Format(CDbl(txtResuMargenBrutoCaja.Text) / CDbl(txtIngresos.Text), "##,#0.00")
    Case 2:
            txtMonParalelo.Text = Format((CDbl(txtIncIngreso.Text) * CDbl(txtEstMonIngreso.Text) * CDbl(txtMagBruto.Text)) - CDbl(txtEstMonOtrosGasto.Text) - CDbl(txtEstMonConsFamiliar.Text) - CDbl(txtCutCredVigente.Text) + CDbl(txtEstMonOtrosIngresos.Text), "##,#0.00")
    Case 3:
            txtIncIngreso.Text = (1 + (Format(CDbl(spnDatosIncrIngreso.valor), "##,#0.00") / 100))
    Case 4:
            If txtMonPropuesto.Text < txtMonParalelo Then
               
               Else
                    MsgBox "El Monto Propuesto es Mayor Al Monto Paralelo ", vbInformation, "Aviso"
                    
                    cmdGuardar.Enabled = False
                
            End If
    End Select
    Exit Sub
    
ErrorCalculo:
MsgBox "Error: Ingrese los datos Correctamente." & Chr(13) & "Detalles de error: " & Err.Description, vbCritical, "Error"

Select Case pnTipo
    Case 1:
            txtResuMargenBrutoCaja.Text = "0.00"
            txtIngresos.Text = "0.00"
    Case 2:
            txtIngresos.Text = "0.00"
            txtResumenIncIngresos.Text = "0.00"
            txtMagBruto.Text = "0.00"
            txtEstMonOtrosGasto.Text = "0.00"
            txtEstMonConsFamiliar.Text = "0.00"
            txtCutCredVigente.Text = "0.00"
            txtEstMonOtrosIngresos.Text = "0.00"
    Case 3:
            spnDatosIncrIngreso.valor = "0.00"
    
End Select
Call CalculoTotal(pnTipo)
     
End Sub

Private Sub Form_Activate()
   txtVentas.SetFocus
End Sub

Private Sub txtFechaVista_LostFocus()

If Not IsDate(txtFechaVista) Then
    MsgBox "Verifique Dia,Mes,Año , Fecha Incorrecta", vbInformation, "Aviso"
    txtFechaVista.SetFocus
End If

End Sub

Private Sub txtVentas_LostFocus()

If Len(Trim(txtVentas.Text)) = "" Then
        txtVentas.Text = "0.00"
    End If
        txtVentas.Text = Format(txtVentas.Text, "#,##0.00")
    
    txtEstMonIngreso.Text = txtVentas.Text
    txtIngresos.Text = txtVentas.Text
    
    Call CalculoTotal(2)
    
End Sub

Private Sub txtCapPago_LostFocus()
    If Len(Trim(txtCapPago.Text)) = "" Then
        txtCapPago.Text = "0.00"
    End If
        txtCapPago.Text = Format(txtCapPago.Text, "#,##0.00")
        
        Call CalculoTotal(2)
        
End Sub

Private Sub txtIngNeto_LostFocus()
If Len(Trim(txtIngNeto.Text)) = "" Then
        txtIngNeto.Text = "0.00"
    End If
        txtIngNeto.Text = Format(txtIngNeto.Text, "#,##0.00")
        
        Call CalculoTotal(2)
        
End Sub
