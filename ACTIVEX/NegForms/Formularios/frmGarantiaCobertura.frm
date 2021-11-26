VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGarantiaCobertura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Cobertura"
   ClientHeight    =   10455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   Icon            =   "frmGarantiaCobertura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10455
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAutoCobertura 
      Caption         =   "Autorización  Cobertura"
      Height          =   375
      Left            =   4200
      TabIndex        =   30
      ToolTipText     =   "Solicitar Exoneración Cobertura"
      Top             =   9780
      Width           =   1845
   End
   Begin VB.TextBox txtTipoCambio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   10690
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdExoneracionCobertura 
      Caption         =   "Exoneración Cobertura"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   25
      ToolTipText     =   "Solicitar Exoneración Cobertura"
      Top             =   9780
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1105
      TabIndex        =   24
      ToolTipText     =   "Eliminar Cobertura"
      Top             =   9780
      Width           =   1000
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10760
      TabIndex        =   19
      ToolTipText     =   "Salir"
      Top             =   9780
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9720
      TabIndex        =   18
      ToolTipText     =   "Cancelar"
      Top             =   9780
      Width           =   1000
   End
   Begin VB.CommandButton cmdSolicitud 
      Caption         =   "&Solicitud"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2140
      TabIndex        =   17
      ToolTipText     =   "Ir a Solicitud"
      Top             =   9780
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   80
      TabIndex        =   16
      ToolTipText     =   "Grabar Cobertura"
      Top             =   9780
      Width           =   1000
   End
   Begin VB.CommandButton cmdSugerencia 
      Caption         =   "Su&gerencia"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3175
      TabIndex        =   15
      ToolTipText     =   "Ir a Sugerencia"
      Top             =   9780
      Width           =   1000
   End
   Begin VB.Frame fraGarantiaCredito 
      Caption         =   "Garantías seleccionadas en el crédito"
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
      Height          =   3600
      Left            =   80
      TabIndex        =   13
      Top             =   6120
      Width           =   11655
      Begin VB.TextBox txtMontoCoberturaTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   350
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   3165
         Width           =   1695
      End
      Begin NegForms.FlexEdit feGarantiaCredito 
         Height          =   2880
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   5080
         Cols0           =   17
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   $"frmGarantiaCobertura.frx":030A
         EncabezadosAnchos=   "400-0-2500-1000-1400-1200-800-1200-1200-0-0-0-0-0-0-1300-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-8-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-L-R-R-R-R-C-C-C-C-R-R-R-C"
         FormatosEdit    =   "0-0-0-0-0-2-2-2-2-0-0-0-0-2-2-2-0"
         CantEntero      =   12
         AvanceCeldas    =   1
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Monto Cobertura TC:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   7995
         TabIndex        =   21
         Top             =   3230
         Width           =   1740
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Index           =   0
         Left            =   7875
         Top             =   3165
         Width           =   3645
      End
   End
   Begin VB.Frame frmGarantiaPersona 
      Caption         =   "Garantías a utilizar"
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
      Height          =   3360
      Left            =   80
      TabIndex        =   11
      Top             =   2700
      Width           =   11655
      Begin VB.CheckBox chkGarantiaPersona 
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   280
         Width           =   210
      End
      Begin NegForms.FlexEdit feGarantiaPersona 
         Height          =   3000
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   5292
         Cols0           =   22
         FixedCols       =   2
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   $"frmGarantiaCobertura.frx":03A6
         EncabezadosAnchos=   "0-0-400-2500-1000-1400-1200-1200-800-800-1200-1200-0-0-0-0-0-0-0-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-L-L-R-R-R-R-L-L-C-C-C-C-C-R-L-C-C-R"
         FormatosEdit    =   "0-0-0-0-0-0-2-2-0-0-0-0-0-0-0-0-0-2-0-0-0-2"
         CantEntero      =   12
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
   End
   Begin VB.Frame fraPersona 
      Caption         =   "Personas relacionadas al crédito"
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
      Height          =   2160
      Left            =   80
      TabIndex        =   2
      Top             =   500
      Width           =   11655
      Begin VB.TextBox txtClienteTipo 
         Height          =   285
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtDestino 
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtMoneda 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtProductoDesc 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2895
      End
      Begin NegForms.FlexEdit fePersona 
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   2143
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "#-Codigo-Nombre del Cliente-Relación-N° DOI"
         EncabezadosAnchos=   "400-0-4500-2200-1500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3-X"
         ListaControles  =   "0-0-0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de Cliente:"
         Height          =   255
         Left            =   8880
         TabIndex        =   28
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Destino :"
         Height          =   255
         Left            =   5640
         TabIndex        =   9
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Monto :"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   1470
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Producto :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1470
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdExaminar 
      Caption         =   "E&xaminar"
      Height          =   375
      Left            =   3825
      TabIndex        =   1
      ToolTipText     =   "Examinar"
      Top             =   80
      Width           =   900
   End
   Begin NegForms.ActXCodCta ActxCuenta 
      Height          =   420
      Left            =   80
      TabIndex        =   0
      Top             =   80
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   741
      Texto           =   " Crédito :"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
      CMAC            =   "109"
   End
   Begin MSComctlLib.StatusBar sbMensaje 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   26
      Top             =   10185
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8467
            MinWidth        =   8467
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Cambio:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   9360
      TabIndex        =   23
      Top             =   160
      Width           =   1080
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   1
      Left            =   9240
      Top             =   120
      Width           =   2430
   End
End
Attribute VB_Name = "frmGarantiaCobertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************************
'** Nombre : frmGarantiaCobertura
'** Descripción : Para configuración del gravamen para los productos creado segun TI-ERS063-2014
'** Creación : EJVG, 20150324 10:00:00 AM
'***********************************************************************************************
Option Explicit

Enum eInicioGravamen
    InicioGravamenxSolicitud = 1
    InicioGravamenxMenu = 2
    InicioGravamenxAjuste = 3
End Enum
Enum eProductoGravamen
    Credito = 1
    CartaFianza = 2
End Enum

Dim fnInicio As eInicioGravamen
Dim fnProducto As eProductoGravamen

Dim fnMoneda As Moneda
Dim fsCtaCod As String
Dim fsPersCodTit As String
Dim fbTpoProdCodCambia As Boolean
Dim fsTpoProdCod As String
Dim fsTpoProdDesc As String
Dim fnMonto As Currency
Dim fnPrdEstado As ColocEstado
Dim fvGravamen() As tGarantiaGravamen

Dim fbCliPreferencial As Boolean
Dim fnTipoCamb As Currency

Dim fnExoneraID As Long
Dim fnExoneraAprueba As Integer '-1: Pendiente, 0: Desaprueba, 1: Aprueba
Dim fnExoneraTasa As Currency

Dim fbAceptar As Boolean
Dim fbSalir As Boolean
Dim fbFocoGrilla As Boolean
Dim fbCheckGrilla As Boolean
Dim fbLeasing As Boolean
Dim fbDataGarantiaCredito As Boolean
Dim fbAmpliaRefinancia As Boolean 'EJVG20160401

Public Function Inicio(ByVal pnInicio As eInicioGravamen, Optional ByVal pnProducto As eProductoGravamen = Credito, _
                    Optional ByVal psCtaCod As String = "", Optional ByVal pbLeasing As Boolean = False, _
                    Optional ByVal pbTpoProdCodCambia As Boolean = False, Optional ByVal psTpoProdCod As String = "", Optional ByVal psTpoProdDesc As String = "", _
                    Optional ByVal pnMonto As Currency = 0, Optional ByRef pvGravamen As Variant = Nothing) As Boolean
                        
    fnInicio = pnInicio
    fsCtaCod = psCtaCod
    fnProducto = pnProducto
    fbLeasing = pbLeasing
    fbTpoProdCodCambia = pbTpoProdCodCambia
    fsTpoProdCod = psTpoProdCod
    fsTpoProdDesc = UCase(Trim(psTpoProdDesc))
    fnMonto = pnMonto
    If IsArray(pvGravamen) Then 'EJVG20160321
        fvGravamen = pvGravamen
    End If
    cmdExoneracionCobertura.Visible = False 'RECO20160628 ERS002-2016
    Show 1
    pvGravamen = fvGravamen
    Inicio = fbAceptar
End Function
'RECO20160628 ERS002-2016 ******************************************************************
Private Sub cmdAutoCobertura_Click()
    Dim oCredNiv As New COMDCredito.DCOMNivelAprobacion
    Dim oCred As New COMDCredito.DCOMCredito
    Dim oRS As New ADODB.Recordset
    Dim sMovNro As String
    Dim lsRatio As String
    Dim lbOk As Boolean
    Dim lnRatio As Double, lnRatio_NEW As Double
    Dim lnRatioConstSistema As Double 'FRHU 20160811 ANEXO-002 ERS002-2016
    
    lnRatioConstSistema = CDbl(Trim(LeeConstanteSist(gConstSistMinimoRatioCobertura))) 'FRHU 20160811 ANEXO-002 ERS002-2016
    If feGarantiaCredito.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. primero debe seleccionar solo y unicamente las garantías que se van a utilizar para está operación.", vbInformation, "Aviso"
        cmdExoneracionCobertura.Enabled = True
        Exit Sub
    End If
    
    lnRatio = CCur(feGarantiaCredito.TextMatrix(1, 6))
    
    Set oRS = oCredNiv.ObtieneDatosNivelAutoCta(ActxCuenta.NroCuenta, "TIP0007")
    If Not (oRS.EOF And oRS.BOF) Then
        If oRS!cMovNroAuto = "" Then
            sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            
            While (Not lbOk)
                lsRatio = Trim(InputBox("Usuario: " & gsCodUser & Chr(13) & "Agencia: " & UCase(gsNomAge) & Chr(13) & "Fecha " & Format(gdFecSis, gsFormatoFechaView) & Chr(13) & Chr(13) & "Ratio Cobertura: " & Format(lnRatio, "#0.0000") & Chr(13) & Chr(13) & "Ingrese el Ratio de Cobertura deseado:", "Autorización por Cobertura Crédito N° " & ActxCuenta.NroCuenta))
                
                If Trim(lsRatio) = "" Then
                    Exit Sub
                End If
                
                If Not IsNumeric(lsRatio) Then
                    MsgBox "Ingrese un valor numerico"
                Else
                    lnRatio_NEW = CDbl(lsRatio)
                                    
                    If lnRatio_NEW >= lnRatio Then
                        MsgBox "Ud. primero debe ingresar un ratio menor a " & Format(lnRatio, "#0.0000"), vbInformation, "Aviso"
                        'cmdExoneracionCobertura.Enabled = True
                        
                        'Exit Sub
                    'FRHU 20160811 ANEXO-002 ERS002-2016
                    ElseIf lnRatio_NEW < lnRatioConstSistema Then
                        MsgBox "Ud. no puede ingresar un ratio menoar a " & Format(lnRatioConstSistema, "#0.0000"), vbInformation, "Aviso"
                    'FIN FRHU 20160811
                    Else
                        lbOk = True
                    End If
                End If
            Wend
            
            MsgBox "Se agregará una autorización por 'Cobertura de Garantías Inscritas a favor de la Caja Maynas' ", vbInformation, "Alerta"
            Call oCredNiv.RegistroAutorizacionManual(ActxCuenta.NroCuenta, oRS!cAutorizaCod, oRS!cNivAprCod, EstadoAutoExonera.gEstadoPendiente, "", sMovNro, oRS!nPrdEstado)
            Call oCred.ActualizaRatioCoberturaAnalista(ActxCuenta.NroCuenta, lnRatio_NEW)
            Call CargaDatos(ActxCuenta.NroCuenta)
        Else
            MsgBox "Ya existe una solicitud pendiente", vbInformation, "Alerta"
        End If
    Else
        MsgBox "No se encuentra datos correspondientes a los niveles de aprobación", vbInformation, "Alerta"
    End If
End Sub
'RECO FIN ***********************************************************************************
Private Sub CmdEliminar_Click()
    Dim obj As COMNCredito.NCOMCredito
    Dim objPista As COMManejador.Pista
    Dim nNroFilas As Integer
    
    On Error GoTo ErrEliminar
    nNroFilas = Val(feGarantiaCredito.TextMatrix(feGarantiaCredito.Rows - 1, 0))
    
    cmdEliminar.Enabled = False
    If nNroFilas <= 0 Then
        MsgBox "No existen garantías que coberturen el crédito", vbInformation, "Aviso"
        cmdEliminar.Enabled = True
        Exit Sub
    End If
    
    If fnInicio <> InicioGravamenxAjuste Then 'Guardamos en memoria para posterior uso
        If MsgBox("¿Está seguro de eliminar la actual cobetura de la" & IIf(nNroFilas = 1, "", "s") & " garantía" & IIf(nNroFilas = 1, "", "s") & " con el crédito?", vbInformation + vbYesNo, "Aviso") = vbNo Then
            cmdEliminar.Enabled = True
            Exit Sub
        End If
        
        Screen.MousePointer = 11
        
        Set obj = New COMNCredito.NCOMCredito
        obj.EliminarCoberturaGarantia ActxCuenta.NroCuenta
        Set obj = Nothing
        
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gEliminar, , ActxCuenta.NroCuenta, gCodigoCuenta
        Set objPista = Nothing
        
        fbDataGarantiaCredito = False
        'DeterminaAccionExoneracion 'RECO20160628 ERS002-2016
        Screen.MousePointer = 0
    End If
    
    cmdEliminar.Enabled = False
    'FormateaFlex feGarantiaCredito
    chkGarantiaPersona.value = 0 'Quitamos los checks actuales
    chkGarantiaPersona_Click
    
    MsgBox "Se ha eliminado la cobertura de la" & IIf(nNroFilas = 1, "", "s") & " garantía" & IIf(nNroFilas = 1, "", "s") & " con el crédito", vbInformation, "Aviso"
    
    Exit Sub
ErrEliminar:
    cmdEliminar.Enabled = True
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdExaminar_Click()
    On Error GoTo ErrorCmdExaminar_Click
    Screen.MousePointer = 11
    If fnProducto = Credito Then
        'ActxCuenta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstSolic, gColocEstSug), "Creditos Solicitados", , , , Right(gsCodAge, 2)) 'COmento JOEP20190208 CP
        'ActxCuenta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstSolic, gColocEstSug), "Creditos Solicitados", , , , Right(gsCodAge, 2), , , , , gsCodCargo) 'JOEP20190208 CP
        If Len(ActxCuenta.NroCuenta) = 18 Then
            Call ActxCuenta_KeyPress(13)
        Else
            Call cmdCancelar_Click
        End If
    ElseIf fnProducto = CartaFianza Then
        'ActxCuenta.NroCuenta = frmCFPersEstado.Inicio(Array(gColocEstSolic, gColocEstSug), "Cartas Fianza")
        If Len(ActxCuenta.NroCuenta) = 18 Then
            Call ActxCuenta_KeyPress(13)
        Else
            Call cmdCancelar_Click
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrorCmdExaminar_Click:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub ActxCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(ActxCuenta.NroCuenta) <> 18 Then
            MsgBox "Credito No Existe", vbInformation, "Aviso"
            Call cmdCancelar_Click
            Exit Sub
        End If
        If Not CargaDatos(ActxCuenta.NroCuenta) Then
            MsgBox "Credito no puede ser seleccionado o no existe", vbInformation, "Aviso"
            Call cmdCancelar_Click
        Else
            fbSalir = False 'EJVG20160318 -> Para que no se cierre el formulario, ya que por esta opción si o si es x digitación de créditos
            If (fnExoneraID > 0 And fnExoneraAprueba = -1) Then
                Call cmdCancelar_Click
                Exit Sub
            End If
            HabilitarBusqueda False
        End If
    End If
End Sub
Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
    Dim oNGar As New COMNCredito.NCOMGarantia
    Dim oDGar As COMDCredito.DCOMGarantia
    Dim rsPersonas As ADODB.Recordset
    Dim rsDatos As ADODB.Recordset
    Dim rsGarantiasPersona As ADODB.Recordset
    Dim rsGarantiasCredito As ADODB.Recordset
    Dim rsColocGarantiaAnt As ADODB.Recordset
    Dim Index As Integer
    Dim i As Integer
    Dim lsColor As Long
    Dim bPermiteModificarCobertura As Boolean
    Dim lsCadenaPermiteModificarCobertura As String
    'Dim lsCadenaGarantiasMigrar As String
    Dim lbBloqueaGarantiaPersona As Boolean
    Dim lbMoneyEqual As Boolean
    Dim iBusca As Integer
    
    On Error GoTo ErrorCargaDatos
    
    MsgBox "Se está cargando los datos, espere un momento..", vbInformation, "Aviso"
    
    fbSalir = False
    ReDim fvColocGarantiaAnt(2, 0)
    
    Screen.MousePointer = 11
    'Call oNGar.CargarDatosxGravamen(psCtaCod, IIf(fnProducto = CartaFianza, True, False), gdFecSis, fsPersCodTit, fbCliPreferencial, rsDatos, rsPersonas, rsGarantiasPersona, rsGarantiasCredito, IIf(fnInicio = InicioGravamenxAjuste, fsTpoProdCod, ""), lsCadenaGarantiasMigrar, rsColocGarantiaAnt)
    Call oNGar.CargarDatosxGravamen(psCtaCod, IIf(fnProducto = CartaFianza, True, False), gdFecSis, fsPersCodTit, fbCliPreferencial, rsDatos, rsPersonas, rsGarantiasPersona, rsGarantiasCredito, IIf(fnInicio = InicioGravamenxAjuste, fsTpoProdCod, ""), rsColocGarantiaAnt)
    Set oNGar = Nothing
    
    If rsDatos.EOF Then Exit Function
        
    txtProductoDesc.Text = IIf(fnInicio = InicioGravamenxAjuste, fsTpoProdDesc, rsDatos!cProducto)
    txtMoneda.Text = rsDatos!cSimbolo
    txtMonto.Text = Format(IIf(fnInicio = InicioGravamenxAjuste, fnMonto, rsDatos!nMonto), gsFormatoNumeroView)
    txtDestino.Text = rsDatos!cDestino
    txtClienteTipo.Text = IIf(fbCliPreferencial, "PREFERENCIAL", "NO PREFERENCIAL")
    fsTpoProdCod = IIf(fnInicio = InicioGravamenxAjuste, fsTpoProdCod, rsDatos!cTpoProdCod)
    fnMoneda = rsDatos!nMoneda
    fnPrdEstado = rsDatos!nPrdEstado
    
    lsColor = IIf(fnMoneda = gMonedaNacional, &HC0FFFF, &HC0FFC0)
    txtProductoDesc.BackColor = lsColor
    txtMoneda.BackColor = lsColor
    txtMonto.BackColor = lsColor
    txtDestino.BackColor = lsColor
    txtClienteTipo.BackColor = lsColor
    
    'fnExoneraID = rsDatos!nExoneraCoberturaID
    'fnExoneraAprueba = rsDatos!nExoneraCoberturaAutoriza
    'fnExoneraTasa = rsDatos!nExoneraCoberturaTasa
    fnExoneraTasa = rsDatos!nRatioCoberturaAnal 'RECO20160713 ERS002-2016
    
    RSClose rsDatos
    
    'FormateaFlex fePersona
    Do While Not rsPersonas.EOF
        fePersona.AdicionaFila
        Index = fePersona.row
        fePersona.TextMatrix(Index, 1) = rsPersonas!cPersCod
        fePersona.TextMatrix(Index, 2) = rsPersonas!cPersNombre
        fePersona.TextMatrix(Index, 3) = rsPersonas!cRelacion
        fePersona.TextMatrix(Index, 4) = rsPersonas!cDOI
        rsPersonas.MoveNext
    Loop
    RSClose rsPersonas
    fePersona.row = 1
    fePersona.TopRow = 1
    
    'FormateaFlex feGarantiaPersona
    'No mostrar columnas en Próducto RapiFlash
    feGarantiaPersona.ColWidth(8) = 795
    feGarantiaPersona.ColWidth(9) = 795
    If fsTpoProdCod = "703" Then
        feGarantiaPersona.ColWidth(8) = 0
        feGarantiaPersona.ColWidth(9) = 0
    End If
    
    If Not rsGarantiasPersona.EOF Then
        Do While Not rsGarantiasPersona.EOF
            feGarantiaPersona.AdicionaFila
            Index = feGarantiaPersona.row
            Call PintarFilaGarantiaPersona(Index, rsGarantiasPersona)
            rsGarantiasPersona.MoveNext
        Loop
    End If
    RSClose rsGarantiasPersona
    feGarantiaPersona.row = 1
    feGarantiaPersona.TopRow = 1
    
    FormateaFlex feGarantiaCredito
    'No mostrar columnas en Próducto RapiFlash
    feGarantiaCredito.ColWidth(6) = 795
    If fsTpoProdCod = "703" Then
        feGarantiaCredito.ColWidth(6) = 0
    End If
    
    lbMoneyEqual = True
    fbDataGarantiaCredito = False
    If Not rsGarantiasCredito.EOF Then
        Do While Not rsGarantiasCredito.EOF
            feGarantiaCredito.AdicionaFila
            Index = feGarantiaCredito.row
            feGarantiaCredito.TextMatrix(Index, 1) = rsGarantiasCredito!cNumGarant 'ID
            feGarantiaCredito.TextMatrix(Index, 2) = rsGarantiasCredito!cPersNombre 'Titular
            feGarantiaCredito.TextMatrix(Index, 3) = rsGarantiasCredito!cNumGarant 'Garantia
            feGarantiaCredito.TextMatrix(Index, 4) = rsGarantiasCredito!cBienGarantia 'Bien
            feGarantiaCredito.TextMatrix(Index, 5) = Format(rsGarantiasCredito!nSaldoCobertura, gsFormatoNumeroView) 'Saldo Cobertura
            feGarantiaCredito.TextMatrix(Index, 6) = Format(rsGarantiasCredito!nRatio, "0.00##############") 'Ratio
            feGarantiaCredito.TextMatrix(Index, 7) = Format(rsGarantiasCredito!nCoberturaMax, gsFormatoNumeroView) 'Max Cob
            feGarantiaCredito.TextMatrix(Index, 8) = Format(rsGarantiasCredito!nGravado, gsFormatoNumeroView) 'Monto Cob
            feGarantiaCredito.TextMatrix(Index, 9) = rsGarantiasCredito!nOrden 'Orden
            feGarantiaCredito.TextMatrix(Index, 10) = rsGarantiasCredito!nMoneda 'Moneda
            If rsGarantiasCredito!nMoneda = Moneda.gMonedaNacional Then
                feGarantiaCredito.BackColorRow &HC0FFFF 'vbYellow
            ElseIf rsGarantiasCredito!nMoneda = Moneda.gMonedaExtranjera Then
                feGarantiaCredito.BackColorRow &HC0FFC0 'vbGreen
            End If
            feGarantiaCredito.TextMatrix(Index, 11) = rsGarantiasCredito!dValorizacion 'Fecha Valorizacion
            feGarantiaCredito.TextMatrix(Index, 12) = rsGarantiasCredito!dTramiteLegal 'Fecha Tramite Legal
            
            'Checkea las Garantias de las personas
            If feGarantiaPersona.TextMatrix(1, 0) <> "" Then
                For i = 1 To feGarantiaPersona.Rows - 1
                    If feGarantiaPersona.TextMatrix(i, 1) = feGarantiaCredito.TextMatrix(Index, 1) Then
                        feGarantiaPersona.TextMatrix(i, 2) = "1"
                        feGarantiaCredito.TextMatrix(Index, 13) = i 'Index referencia Garantia-Persona
                        Exit For
                    End If
                Next
            End If
            feGarantiaCredito.TextMatrix(Index, 14) = Format(rsGarantiasCredito!nGravado, gsFormatoNumeroView) 'Monto Cob Temporal
            'feGarantiaCredito.TextMatrix(Index, 15) = Format(Round(rsGarantiasCredito!nGravado * IIf(rsGarantiasCredito!nMoneda = Moneda.gMonedaExtranjera, CCur(txtTipoCambio.Text), 1), 2), gsFormatoNumeroView) 'Monto Cob Tipo Cambio
            If fnMoneda <> CInt(feGarantiaCredito.TextMatrix(Index, 10)) Then 'Comparamos moneda del crédito con la garantia
                If fnMoneda = gMonedaNacional Then 'Garantia Dolares -> Crédito Soles
                    feGarantiaCredito.TextMatrix(Index, 15) = Format(Round(CCur(feGarantiaCredito.TextMatrix(Index, 8)) * CCur(txtTipoCambio.Text), 2), gsFormatoNumeroView)
                ElseIf fnMoneda = gMonedaExtranjera Then 'Garantia Soles -> Crédito Dolares
                    feGarantiaCredito.TextMatrix(Index, 15) = Format(Round(CCur(feGarantiaCredito.TextMatrix(Index, 8)) / CCur(txtTipoCambio.Text), 2), gsFormatoNumeroView)
                End If
            Else
                feGarantiaCredito.TextMatrix(Index, 15) = Format(CCur(feGarantiaCredito.TextMatrix(Index, 8)), gsFormatoNumeroView)
            End If

            If fnMoneda <> CInt(feGarantiaCredito.TextMatrix(Index, 10)) Then
                lbMoneyEqual = False
            End If
            
            rsGarantiasCredito.MoveNext
        Loop
        fbDataGarantiaCredito = True
    End If
    RSClose rsGarantiasCredito
    
    'Refinanciaciones y Ampliaciones (Inicialmente chekeadas con las Garantías de los créditos Anteriores)
    fbAmpliaRefinancia = False
    If Not fbDataGarantiaCredito Then
        If Not rsColocGarantiaAnt.EOF Then
            fbAmpliaRefinancia = True
            'Damos check en la lista de garantías actual de los intervinientes
            Do While Not rsColocGarantiaAnt.EOF
                For iBusca = 1 To feGarantiaPersona.Rows - 1
                    'If feGarantiaPersona.TextMatrix(iBusca, 1) = rsColocGarantiaAnt!cNumGarant Then
                    If feGarantiaPersona.TextMatrix(iBusca, 1) = rsColocGarantiaAnt!cNumGarant And Not CBool(rsColocGarantiaAnt!bMigrado) Then
                        feGarantiaPersona.TextMatrix(iBusca, 2) = "1"
                        Exit For
                    End If
                Next
                rsColocGarantiaAnt.MoveNext
            Loop
            'Hacemos el recalculo para todos las garantías chekeadas
            feGarantiaPersona_OnCellCheck 0, 0
        End If
    End If
    RSClose rsColocGarantiaAnt
    
    feGarantiaCredito.row = 1
    feGarantiaCredito.TopRow = 1
    
    If fnInicio = InicioGravamenxAjuste Then
        'Si es x Ajuste y se ha modificado el Tipo de Producto se limpiará todo el Gravamen, ya que puede que los ratios varían según nuevo tipo de Producto
        If fbTpoProdCodCambia Then
            chkGarantiaPersona.value = 0 'Quitamos los checks actuales
            chkGarantiaPersona_Click
        End If
    End If
    
    txtMontoCoberturaTC.Text = Format(feGarantiaCredito.SumaRow(15), gsFormatoNumeroView)

    cmdGrabar.Enabled = False
    cmdEliminar.Enabled = False
    
    'Recupera cadena que permita modificar Cobertura
    Set oDGar = New COMDCredito.DCOMGarantia
    lsCadenaPermiteModificarCobertura = oDGar.CadenaPermiteModificarCobertura(psCtaCod, lbMoneyEqual, CCur(txtMonto.Text), CCur(txtMontoCoberturaTC.Text), lbBloqueaGarantiaPersona)
    Set oDGar = Nothing
    
    If Len(lsCadenaPermiteModificarCobertura) > 0 Then
        fbSalir = True
        'If Len(lsCadenaGarantiasMigrar) = 0 Then 'Para mostrar solo el último mensaje
            MsgBox lsCadenaPermiteModificarCobertura, vbInformation, "Aviso"
        'End If
    End If
    
    cmdGrabar.Enabled = IIf(Len(lsCadenaPermiteModificarCobertura) = 0 And feGarantiaPersona.TextMatrix(1, 0) <> "", True, False)
    cmdEliminar.Enabled = IIf(Len(lsCadenaPermiteModificarCobertura) = 0 And feGarantiaCredito.TextMatrix(1, 0) <> "", True, False)
    
    frmGarantiaPersona.Enabled = Not lbBloqueaGarantiaPersona
    If lbBloqueaGarantiaPersona Then
        cmdEliminar.Enabled = Not lbBloqueaGarantiaPersona
    End If
            
    Screen.MousePointer = 0
            
    'Exoneraciones
    If fbSalir Then
        'cmdExoneracionCobertura.Enabled = False 'RECO20160628 ERS002-2016
    Else
        'DeterminaAccionExoneracion True, fbSalir 'RECO20160628 ERS002-2016
    End If

    CargaDatos = True
    Exit Function
ErrorCargaDatos:
    CargaDatos = False
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
Private Sub PintarFilaGarantiaPersona(ByVal pnIndex As Integer, ByVal prsGarantiasPersona As ADODB.Recordset)
    feGarantiaPersona.TextMatrix(pnIndex, 1) = prsGarantiasPersona!cNumGarant 'ID
    feGarantiaPersona.TextMatrix(pnIndex, 3) = prsGarantiasPersona!cPersNombreTit 'Nombre Tit
    feGarantiaPersona.TextMatrix(pnIndex, 4) = prsGarantiasPersona!cNumGarant 'Clasificacion
    feGarantiaPersona.TextMatrix(pnIndex, 5) = prsGarantiasPersona!cBienGarantia 'Bien en Garantía
    feGarantiaPersona.TextMatrix(pnIndex, 6) = Format(prsGarantiasPersona!nValorGarantia, gsFormatoNumeroView) 'Valor de Garantía
    feGarantiaPersona.TextMatrix(pnIndex, 7) = Format(prsGarantiasPersona!nDisponible, gsFormatoNumeroView) 'Monto Disponible sin considerar la misma cobertura, ni la cobertura de los créditos a ampliar o los créditos a refinanciar
    feGarantiaPersona.TextMatrix(pnIndex, 8) = Format(prsGarantiasPersona!nRatioP, "0.00##############") 'Ratio Preferida
    feGarantiaPersona.TextMatrix(pnIndex, 9) = Format(prsGarantiasPersona!nRatioNP, "0.00##############") 'Ratio No Preferida
    feGarantiaPersona.TextMatrix(pnIndex, 10) = prsGarantiasPersona!cPersNombreEmi 'Nombre Emisor
    feGarantiaPersona.TextMatrix(pnIndex, 11) = prsGarantiasPersona!cDocNro 'Nro. Doc.
    feGarantiaPersona.TextMatrix(pnIndex, 12) = IIf(prsGarantiasPersona!bPreferida, 1, 0) 'Preferida
    feGarantiaPersona.TextMatrix(pnIndex, 13) = prsGarantiasPersona!nOrden 'Orden
    feGarantiaPersona.TextMatrix(pnIndex, 14) = prsGarantiasPersona!nMoneda 'Moneda
    If prsGarantiasPersona!nMoneda = Moneda.gMonedaNacional Then
        feGarantiaPersona.BackColorRow &HC0FFFF 'vbYellow
    ElseIf prsGarantiasPersona!nMoneda = Moneda.gMonedaExtranjera Then
        feGarantiaPersona.BackColorRow &HC0FFC0 'vbGreen
    End If
    feGarantiaPersona.TextMatrix(pnIndex, 15) = prsGarantiasPersona!dValorizacion 'FechaValorizacion
    feGarantiaPersona.TextMatrix(pnIndex, 16) = prsGarantiasPersona!dTramiteLegal 'FechaTramiteLegal
    feGarantiaPersona.TextMatrix(pnIndex, 17) = prsGarantiasPersona!nMontoDisponibleAL 'Monto Disponible de ultima valorización AutoLiquidable
    feGarantiaPersona.TextMatrix(pnIndex, 18) = IIf(prsGarantiasPersona!bVAL_Migra, 1, 0) 'Indica si última valorización fue migrada
    feGarantiaPersona.TextMatrix(pnIndex, 19) = IIf(prsGarantiasPersona!bTRA_Migra, 1, 0) 'Indica si última valorización fue migrada
    feGarantiaPersona.TextMatrix(pnIndex, 20) = prsGarantiasPersona!nTipoValorizacion 'Tipo de Valorizacion
    feGarantiaPersona.TextMatrix(pnIndex, 21) = Format(prsGarantiasPersona!nDisponibleGAR, gsFormatoNumeroView) 'Monto Disponible Real
End Sub
Private Sub Limpiar()
    ActxCuenta.NroCuenta = ""
    ActxCuenta.CMAC = gsCodCMAC
    ActxCuenta.Age = Right(gsCodAge, 2)
    FormateaFlex fePersona
    txtProductoDesc.Text = ""
    txtProductoDesc.BackColor = &H80000005
    txtMoneda.Text = ""
    txtMoneda.BackColor = &H80000005
    txtMonto.Text = "0.00"
    txtMonto.BackColor = &H80000005
    txtDestino.Text = ""
    txtDestino.BackColor = &H80000005
    txtClienteTipo.Text = ""
    txtClienteTipo.BackColor = &H80000005
    chkGarantiaPersona.value = 0
    FormateaFlex feGarantiaPersona
    FormateaFlex feGarantiaCredito
    txtMontoCoberturaTC.Text = "0.00"
    txtTipoCambio.Text = Format(fnTipoCamb, gsFormatoNumeroView)
    sbMensaje.Panels(1).Text = ""
    cmdExoneracionCobertura.Visible = False 'RECO20160628 ERS002-2016
End Sub
Private Sub LimpiarVariables()
    fbCheckGrilla = False
    'fsCtaCod = ""
    fsPersCodTit = ""
    fbCliPreferencial = False
    fsTpoProdCod = ""
    'fnTipoCamb = 0#
    fbSalir = False
End Sub
Private Sub HabilitarBusqueda(ByVal pbHabilitar As Boolean)
    ActxCuenta.Enabled = pbHabilitar
    CmdExaminar.Enabled = pbHabilitar
End Sub
Private Sub cmdCancelar_Click()
    Limpiar
    LimpiarVariables
    HabilitarBusqueda (True)
    cmdGrabar.Enabled = False
    cmdEliminar.Enabled = False
    'cmdExoneracionCobertura.Enabled = False 'RECO20160628 ERS002-2016
End Sub

Private Sub cmdExoneracionCobertura_Click()
    Dim obj As COMNCredito.NCOMCredito
    Dim oFun As NContFunciones
    Dim lsComentario As String
    Dim lnRatio As Double
    
    On Error GoTo ErrExoneracionTasa
    
    cmdExoneracionCobertura.Enabled = False
    
    If feGarantiaCredito.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. primero debe seleccionar solo y unicamente las garantías que se van a utilizar para está operación.", vbInformation, "Aviso"
        cmdExoneracionCobertura.Enabled = True
        Exit Sub
    End If
    
    lnRatio = CCur(feGarantiaCredito.TextMatrix(1, 6))
    
    If lnRatio <= 1# Then
        MsgBox "El ratio de cobertura actual tiene que ser mayor a uno (1.00)", vbInformation, "Aviso"
        cmdExoneracionCobertura.Enabled = True
        Exit Sub
    End If
    
    lsComentario = Trim(InputBox("Usuario: " & gsCodUser & Chr(13) & "Agencia: " & UCase(gsNomAge) & Chr(13) & "Fecha " & Format(gdFecSis, gsFormatoFechaView) & Chr(13) & Chr(13) & "Ratio Cobertura: " & Format(lnRatio, "#0.0000") & Chr(13) & Chr(13) & "Ingrese descripción y el Ratio de Cobertura esperado:", "Exoneración Cobertura Crédito N° " & ActxCuenta.NroCuenta))
    If Len(lsComentario) <= 0 Then
        cmdExoneracionCobertura.Enabled = True
        Exit Sub
    End If
    
    If MsgBox("Se enviará una solicitud de Exoneración de Cobertura a la Jefatura de Productos Créditicios." & Chr(13) & Chr(13) & "Mientras no se apruebe/rechace la misma no se podrá continuar." & Chr(13) & Chr(13) & "¿Desea continuar?", vbInformation + vbYesNo, "Confirmación") = vbNo Then
        cmdExoneracionCobertura.Enabled = True
        Exit Sub
    End If
    
    Set obj = New COMNCredito.NCOMCredito
    Set oFun = New NContFunciones
    
    fnExoneraID = obj.SolicitudExoneraCobertura(ActxCuenta.NroCuenta, lsComentario, oFun.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), lnRatio)
    fnExoneraAprueba = -1
    fnExoneraTasa = 0#
    
    MsgBox "Se ha enviado la solicitud de Ratio de Cobertura, comuniquese con el encargado para su aprobación/rechazo.", vbInformation, "Aviso"
    
    Set obj = Nothing
    Set oFun = Nothing
    
    DeterminaAccionExoneracion True
    If fnInicio = InicioGravamenxMenu Then
        cmdCancelar_Click
    Else
        Unload Me
    End If
    Exit Sub
ErrExoneracionTasa:
    fnExoneraID = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub DeterminaAccionExoneracion(Optional ByVal pbMsgBox As Boolean = False, Optional ByRef pbPendienteAprobacion As Boolean = False)
    Dim lsmensaje As String
    
    cmdExoneracionCobertura.Visible = True
    cmdExoneracionCobertura.Enabled = True
    
    If fnProducto = CartaFianza Then
        cmdExoneracionCobertura.Visible = False
        Exit Sub
    End If
    If fsTpoProdCod = "703" Then
        cmdExoneracionCobertura.Visible = False
        Exit Sub
    End If
    If fbDataGarantiaCredito Then ' Or (Not fbDataGarantiaCredito And feGarantiaCredito.TextMatrix(1, 0) <> "") Then
        cmdExoneracionCobertura.Enabled = False
        Exit Sub
    End If
    If fnInicio = InicioGravamenxAjuste Then
        cmdExoneracionCobertura.Visible = False
        'Exit Sub
    End If
    
    sbMensaje.Panels(1).Text = ""
    If fnExoneraID > 0 Then
        If fnExoneraAprueba = -1 Then 'Pendiente
            cmdExoneracionCobertura.Enabled = False
            
            lsmensaje = "Está pendiente la Aprobación/Rechazo de la Solicitud de Exoneración de Cobertura"
            If pbMsgBox Then
                MsgBox lsmensaje, vbInformation, "Aviso"
            End If
            pbPendienteAprobacion = True
            sbMensaje.Panels(1).Text = lsmensaje
        ElseIf fnExoneraAprueba = 0 Then 'Desaprueba (Podría solicitar nuevamente otra Tasación)
            cmdExoneracionCobertura.Enabled = True
            
            lsmensaje = "La solicitud de Exoneración de Cobertura fue Rechazada"
            If pbMsgBox Then
                MsgBox lsmensaje, vbInformation, "Aviso"
            End If
            sbMensaje.Panels(1).Text = lsmensaje
        ElseIf fnExoneraAprueba = 1 Then 'Aprueba
            cmdExoneracionCobertura.Enabled = False

            If pbMsgBox Then
                MsgBox "La solicitud de Exoneración de Cobertura fue Aprobada" & Chr(13) & Chr(13) & _
                        "- Cobertura: " & Format(fnExoneraTasa * IIf(fsTpoProdCod = "703", 100#, 1#), "#0.0000") & " %", vbInformation, "Aviso"
            End If
            sbMensaje.Panels(1).Text = "Con Exoneración de Cobertura del " & Format(fnExoneraTasa * IIf(fsTpoProdCod = "703", 100#, 1#), "#0.0000") & " %"
        End If
    End If
End Sub

Private Sub cmdGrabar_Click()
    Dim obj As COMNCredito.NCOMCredito
    Dim objPista As COMManejador.Pista
    Dim bExito As Boolean
    Dim lvDatos() As tGarantiaGravamen
    Dim i As Integer
    Dim lnCuentaOtraIFi As Integer 'EJVG20160227
    Dim lsNumGarantLista As String
    'APRI20170208
    Dim TipoCredito As ADODB.Recordset
    Dim nTipoCredito As Integer
    Dim nTotalGarantiaPF As Integer
    'END APRI
    'CTI5 ERS0012021***************
    Dim lsTpoProdCod As String
    Dim lnTpoBienFuturo As Integer
    Dim lsCtaCod As String
    '******************************
    'JOEP ERS047 20170904
   Dim objGartLiq As COMDCredito.DCOMCredito
   Dim objGartHipo As COMDCredito.DCOMCredito
    'JOEP ERS047 20170904
    
    On Error GoTo ErrGrabar
    cmdGrabar.Enabled = False
    If Not validarGrabar Then
        cmdGrabar.Enabled = True
        Exit Sub
    End If
    
    ReDim lvDatos(0)
    For i = 1 To feGarantiaCredito.Rows - 1
        ReDim Preserve lvDatos(0 To i)
        lvDatos(i).sNumGarant = feGarantiaCredito.TextMatrix(i, 1) 'ID Garantia
        lvDatos(i).nMoneda = feGarantiaCredito.TextMatrix(i, 10) 'Moneda Garantia
        lvDatos(i).dFechaValorizacion = CDate(feGarantiaCredito.TextMatrix(i, 11)) 'Fecha Valorizacion
        lvDatos(i).dFechaTramiteLegal = CDate(feGarantiaCredito.TextMatrix(i, 12)) 'Fecha Tramite Legal
        lvDatos(i).nSaldoCobertura = CDbl(feGarantiaCredito.TextMatrix(i, 5)) 'Saldo Cobertura
        lvDatos(i).nRatio = CDbl(feGarantiaCredito.TextMatrix(i, 6)) 'Ratio
        lvDatos(i).nMaxCobertura = CDbl(feGarantiaCredito.TextMatrix(i, 7)) 'Max. Cobertura
        lvDatos(i).nCobertura = CDbl(feGarantiaCredito.TextMatrix(i, 8)) 'Cobertura
        lvDatos(i).nOrden = CInt(feGarantiaCredito.TextMatrix(i, 9)) 'Orden para liberacion
        
        If Val(feGarantiaPersona.TextMatrix(feGarantiaCredito.TextMatrix(i, 13), 20)) = eGarantiaTipoValorizacion.GravamenFavorOtraIFi Then
            lnCuentaOtraIFi = lnCuentaOtraIFi + 1
        End If
        lsNumGarantLista = lsNumGarantLista & lvDatos(i).sNumGarant & ","
    Next
    
    If UBound(lvDatos) = 0 Then
        MsgBox "Ud. debe seleccionar al menos una garantía para continuar", vbInformation, "Aviso"
        cmdGrabar.Enabled = True
        Exit Sub
    ElseIf UBound(lvDatos) > 1 Then 'EJVG20160227
        If lnCuentaOtraIFi > 0 Then
            If lnCuentaOtraIFi <> UBound(lvDatos) Then
                MsgBox "Si se va a utilizar Garantías con valorización [GRAVAMEN A FAVOR DE OTRA(s) IFI(s)]" & Chr(13) & "todas las garantías deben ser del mismo tipo.", vbInformation, "No se puede continuar"
                cmdGrabar.Enabled = True
                Exit Sub
            End If
        End If
    End If
    
     
    If fnInicio <> InicioGravamenxAjuste Then
        If MsgBox("¿Está seguro de guardar los datos de las coberturas?", vbInformation + vbYesNo, "Aviso") = vbNo Then
            cmdGrabar.Enabled = True
            Exit Sub
        End If
    End If
    
    fbAceptar = True
    
    If fnInicio = InicioGravamenxAjuste Then 'Guardamos en memoria para posterior uso
        fvGravamen = lvDatos
        Unload Me
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Set obj = New COMNCredito.NCOMCredito
    bExito = obj.GrabarCoberturaGarantia(ActxCuenta.NroCuenta, fsTpoProdCod, CCur(txtMonto.Text), lvDatos, , IIf(fbCliPreferencial, 1, 0))
    Set obj = Nothing
    Screen.MousePointer = 0
    
    If Not bExito Then
        MsgBox "Ha sucedido un error al grabar los datos, si el problema persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
        Exit Sub
    End If
    
    If Len(lsNumGarantLista) > 0 Then
        lsNumGarantLista = Mid(lsNumGarantLista, 1, Len(lsNumGarantLista) - 1)
    End If
    
    Set objPista = New COMManejador.Pista
    'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, lsNumGarantLista, ActxCuenta.NroCuenta, gCodigoCuenta 'RECO20161020 ERS060-2016
    'RECO20161020 ERS060-2016 **********************************************************
     Dim oNCOMColocEval As New NCOMColocEval
     Dim lcMovNro As String
     
     If Not ValidaExisteRegProceso(ActxCuenta.NroCuenta, gTpoRegCtrlGarantia) Then
        lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, lsNumGarantLista, ActxCuenta.NroCuenta, gCodigoCuenta
        Call oNCOMColocEval.insEstadosExpediente(ActxCuenta.NroCuenta, "Cobertura de Credito", lcMovNro, "", "", "", 1, 2000, gTpoRegCtrlGarantia)
        Set oNCOMColocEval = Nothing
     End If
     'RECO FIN **************************************************************************
    Set objPista = Nothing
        

        
    MsgBox "Se ha grabado los datos satisfactoriamente", vbInformation, "Aviso"
    
     'APRI 20170206
     Set obj = New COMNCredito.NCOMCredito
     Set TipoCredito = obj.ObtenerTipoCredito(ActxCuenta.NroCuenta)
     Set obj = Nothing
    
    Do While Not TipoCredito.EOF
        nTipoCredito = TipoCredito!cTpoProdCod
        nTotalGarantiaPF = TipoCredito!nTotalGarantiaPF
        TipoCredito.MoveNext
    Loop
    
    If nTipoCredito = 703 And nTotalGarantiaPF > 0 Then
       
        ImprimeCartaAfectacion ActxCuenta.NroCuenta, nTipoCredito, CCur(txtMonto.Text)
  
    End If
    'END APRI
    
    'JOEP Inicio ERS047 20170904
    Set objGartLiq = New COMDCredito.DCOMCredito
    If objGartLiq.VerificaGarantAutoliq(ActxCuenta.NroCuenta, 1) Then
        If objGartLiq.VerificaSuperaUmbralTpCredito(17, CCur(txtMonto.Text), Mid(ActxCuenta.NroCuenta, 9, 1)) Then
            MsgBox "El Crédito supera el porcentaje máximo de Carta Fianza con Garantia Autoliquidable; se podrá continuar, pero no se podrá sugerir si no se tiene la autorización de Riesgos, Desea Continuar?", vbInformation, "Aviso"
            Call objGartLiq.InsertarSolicitudAutorizacionTpCredito(ActxCuenta.NroCuenta, 17, CDbl(txtMonto.Text))
        Else
            Call objGartLiq.EliminarSolicitudAutorizacionZonaxProduxGarant(ActxCuenta.NroCuenta, 17, 2)
        End If
    End If
    Set objGartLiq = Nothing
    '==========
    Set objGartHipo = New COMDCredito.DCOMCredito
    If objGartHipo.VerificaGarantAutoliq(ActxCuenta.NroCuenta, 2) Then
        If objGartHipo.VerificaSuperaUmbralTpCredito(8, CCur(txtMonto.Text), Mid(ActxCuenta.NroCuenta, 9, 1)) Then
            MsgBox "El Crédito supera el porcentaje máximo de Carta Fianza con Garantia Hipotecaria; se podrá continuar, pero no se podrá sugerir si no se tiene la autorización de Riesgos, Desea Continuar?", vbInformation, "Aviso"
            Call objGartHipo.InsertarSolicitudAutorizacionTpCredito(ActxCuenta.NroCuenta, 8, CDbl(txtMonto.Text))
        Else
            Call objGartHipo.EliminarSolicitudAutorizacionZonaxProduxGarant(ActxCuenta.NroCuenta, 8, 2)
        End If
    End If
    Set objGartHipo = Nothing
    'JOEP Fin ERS047 20170904
    
    cmdGrabar.Enabled = True
    cmdEliminar.Enabled = True
    
    fbDataGarantiaCredito = True
    cmdSugerencia.Enabled = fbDataGarantiaCredito
    'DeterminaAccionExoneracion 'RECO20160628 ERS002-2016
    Exit Sub
ErrGrabar:
    cmdGrabar.Enabled = True
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Function validarGrabar() As Boolean
    Dim oCredito As COMDCredito.DCOMCredito
    Dim i As Integer
    Dim lnMontoGarantias As Currency, lnMontoCredito As Currency
    Dim bMayorAprobado As Boolean
    Dim bMoneyEqual As Boolean
    
    If Len(ActxCuenta.NroCuenta) <> 18 Then
        MsgBox "Ud. primero debe especificar el Nro. de Cuenta", vbInformation, "Aviso"
        EnfocaControl ActxCuenta
        Exit Function
    End If
    If feGarantiaCredito.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. primero debe agregar las garantías vinculadas con el Producto", vbInformation, "Aviso"
        Exit Function
    End If
    If Not IsNumeric(txtTipoCambio.Text) Then
        MsgBox "No se ha encontrado el Tipo de Cambio diario", vbInformation, "Aviso"
        Exit Function
    Else
        If CCur(txtTipoCambio.Text) <= 0 Then
            MsgBox "El Tipo de Cambio diario no puede ser menor que cero", vbInformation, "Aviso"
            Exit Function
        End If
    End If
    
    lnMontoGarantias = feGarantiaCredito.SumaRow(15)
    lnMontoCredito = CCur(txtMonto.Text)
    
    bMoneyEqual = True
    For i = 1 To feGarantiaCredito.Rows - 1
        If fnMoneda <> CInt(feGarantiaCredito.TextMatrix(i, 10)) Then 'Comparamos moneda del crédito con la garantia
            bMoneyEqual = False
            Exit For
        End If
    Next
    
    If bMoneyEqual Then 'Cuando sean monedas iguales debe ser exacto la cobertura
        If lnMontoGarantias > lnMontoCredito Then
            MsgBox "El monto total de garantías sobrepasan el préstamo:" & Chr(13) & Chr(13) & "- Préstamo: " & Format(lnMontoCredito, gsFormatoNumeroView) & Chr(13) & "- Garantías: " & Format(lnMontoGarantias, gsFormatoNumeroView), vbInformation, "Aviso"
            Exit Function
        ElseIf lnMontoGarantias < lnMontoCredito Then
            MsgBox "El monto total de garantías no coberturan el préstamo:" & Chr(13) & Chr(13) & "- Préstamo: " & Format(lnMontoCredito, gsFormatoNumeroView) & Chr(13) & "- Garantías: " & Format(lnMontoGarantias, gsFormatoNumeroView), vbInformation, "Aviso"
            Exit Function
        End If
    Else 'Cuando sean monedas diferentes dejaremos pasar diferencia de 0.01 (Basado en el anterior modulo de Gravamen, pero que no pase de S/. 0.20 o US$ 0.20)
        If Not (lnMontoGarantias - lnMontoCredito >= 0# And lnMontoGarantias - lnMontoCredito <= 0.2) Then
            If lnMontoGarantias > lnMontoCredito Then
                MsgBox "El monto total de garantías sobrepasan el préstamo:" & Chr(13) & Chr(13) & "- Préstamo: " & Format(lnMontoCredito, gsFormatoNumeroView) & Chr(13) & "- Garantías: " & Format(lnMontoGarantias, gsFormatoNumeroView), vbInformation, "Aviso"
                Exit Function
            ElseIf lnMontoGarantias < lnMontoCredito Then
                MsgBox "El monto total de garantías no coberturan el préstamo:" & Chr(13) & Chr(13) & "- Préstamo: " & Format(lnMontoCredito, gsFormatoNumeroView) & Chr(13) & "- Garantías: " & Format(lnMontoGarantias, gsFormatoNumeroView), vbInformation, "Aviso"
                Exit Function
            End If
        End If
    End If
    
    For i = 1 To feGarantiaCredito.Rows - 1
        If CCur(feGarantiaCredito.TextMatrix(i, 8)) <= 0 Then
            MsgBox "El monto de cobertura de la garantía N° " & feGarantiaCredito.TextMatrix(i, 1) & " debe ser mayor a cero." & Chr(13) & "Si no se va a utilizar no debe dar check en garantías a utilizar.", vbInformation, "Aviso"
            feGarantiaCredito.row = i
            feGarantiaCredito.TopRow = i
            feGarantiaCredito.Col = 3
            Exit Function
        End If
        If CCur(feGarantiaCredito.TextMatrix(i, 8)) > CCur(feGarantiaCredito.TextMatrix(i, 7)) Then
            MsgBox "El monto de cobertura no puede ser mayor al monto máximo de cobertura", vbInformation, "Aviso"
            feGarantiaCredito.row = i
            feGarantiaCredito.TopRow = i
            feGarantiaCredito.Col = 8
            Exit Function
        End If
        Set oCredito = New COMDCredito.DCOMCredito
        bMayorAprobado = oCredito.GarantiaPerteneceACreditoAprobado(ActxCuenta.NroCuenta, feGarantiaCredito.TextMatrix(i, 1))
        Set oCredito = Nothing
        If bMayorAprobado Then
            MsgBox "Garantia ya esta siendo usada por un crédito vigente, no se puede continuar", vbInformation, "Aviso"
            feGarantiaCredito.row = i
            feGarantiaCredito.TopRow = i
            feGarantiaCredito.Col = 3
            Exit Function
        End If
    Next
    
'JOEP20181220 CP
    If ActxCuenta.Prod = "514" Then
        Dim rsIndetGarantia As ADODB.Recordset
        Dim cNunGart As String
        Set oCredito = New COMDCredito.DCOMCredito
        
        For i = 1 To (feGarantiaCredito.Rows - 1)
            If i = 1 Then
                cNunGart = cNunGart & Trim(feGarantiaCredito.TextMatrix(i, 3))
            Else
                cNunGart = cNunGart & "," & Trim(feGarantiaCredito.TextMatrix(i, 3))
            End If
        Next i
        
        Set rsIndetGarantia = oCredito.CP_IdentGarantiasCF(ActxCuenta.NroCuenta, cNunGart, gsCodCargo)
            If Not (rsIndetGarantia.BOF And rsIndetGarantia.EOF) Then
                If rsIndetGarantia!cMensaje <> "" Then
                    MsgBox rsIndetGarantia!cMensaje, vbInformation, "Aviso"
                    Set oCredito = Nothing
                    RSClose rsIndetGarantia
                    Exit Function
                End If
            End If
        Set oCredito = Nothing
        RSClose rsIndetGarantia
    End If
'JOEP20181220 CP

    VerificarFechaSistema Me, True
    
    validarGrabar = True
End Function
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub cmdSolicitud_Click()
    Unload Me
End Sub
Private Sub cmdSugerencia_Click()
    If fnProducto = Credito Then
        If gnAgenciaCredEval = 0 Then
            '->***** LUCV20180601, Comentó y agregó según ERS022-2018
            'Call frmCredSugerencia.InicioCargaDatos(ActxCuenta.NroCuenta, fbLeasing, True)
            MsgBox "Agencia no configurada para el proceso de la sugerencia, por favor coordinar con TI", vbInformation, "Alerta"
            '<-***** Fin LUCV20180601
        Else
            Call frmCredSugerencia_NEW.InicioCargaDatos(ActxCuenta.NroCuenta, fbLeasing, True)
        End If
    ElseIf fnProducto = CartaFianza Then
        frmCFSugerencia.Inicia (ActxCuenta.NroCuenta)
    End If
End Sub

Private Sub feGarantiaCredito_GotFocus()
    fbFocoGrilla = True
End Sub
Private Sub feGarantiaCredito_LostFocus()
    'fbFcoGrilla = False
End Sub
Private Sub feGarantiaCredito_OnCellChange(pnRow As Long, pnCol As Long)
    Dim iGP As Integer
    If feGarantiaCredito.TextMatrix(pnRow, 0) <> "" Then
        If pnCol = 8 Then
            If Val(feGarantiaCredito.TextMatrix(pnRow, 13)) > 0 Then 'Index referencia de Garantia-Persona
                iGP = Val(feGarantiaCredito.TextMatrix(pnRow, 13))
                'feGarantiaPersona.TextMatrix(iGP, 6) = Format(CCur(feGarantiaPersona.TextMatrix(iGP, 6)) + CCur(feGarantiaCredito.TextMatrix(pnRow, 14)), gsFormatoNumeroView) 'Agrego la cobertura antes de la edición
                'feGarantiaPersona.TextMatrix(iGP, 6) = Format(CCur(feGarantiaPersona.TextMatrix(iGP, 6)) - CCur(feGarantiaCredito.TextMatrix(pnRow, 8)), gsFormatoNumeroView) 'Agrego la cobertura antes de la edición
                feGarantiaCredito.TextMatrix(pnRow, 14) = feGarantiaCredito.TextMatrix(pnRow, 8)  'Actualizamos la columna temporal
                
                'Actualizamos Monto TC (Moneda del crédito)
                If fnMoneda <> CInt(feGarantiaCredito.TextMatrix(pnRow, 10)) Then 'Comparamos moneda del crédito con la garantia
                    If fnMoneda = gMonedaNacional Then 'Garantia Dolares -> Crédito Soles
                        feGarantiaCredito.TextMatrix(pnRow, 15) = Format(Round(CCur(feGarantiaCredito.TextMatrix(pnRow, 8)) * CCur(txtTipoCambio.Text), 2), gsFormatoNumeroView)
                    ElseIf fnMoneda = gMonedaExtranjera Then 'Garantia Soles -> Crédito Dolares
                        feGarantiaCredito.TextMatrix(pnRow, 15) = Format(Round(CCur(feGarantiaCredito.TextMatrix(pnRow, 8)) / CCur(txtTipoCambio.Text), 2), gsFormatoNumeroView)
                    End If
                Else
                    feGarantiaCredito.TextMatrix(pnRow, 15) = Format(CCur(feGarantiaCredito.TextMatrix(pnRow, 8)), gsFormatoNumeroView)
                End If
            End If
            txtMontoCoberturaTC.Text = Format(feGarantiaCredito.SumaRow(15), gsFormatoNumeroView)
        End If
    End If
End Sub
Private Sub feGarantiaCredito_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim bCancel As Boolean
    
    bCancel = EsCeldaFlexEditable(feGarantiaCredito, pnCol)
    If Not bCancel Then
        Cancel = bCancel
        Exit Sub
    End If
    If feGarantiaCredito.TextMatrix(pnRow, 0) <> "" Then
        If pnCol = 8 Then
            If Not IsNumeric(feGarantiaCredito.TextMatrix(pnRow, 8)) Then
                MsgBox "El monto de cobertura debe ser mayor a cero", vbInformation, "Aviso"
                Cancel = False
                Exit Sub
            End If
            If CCur(feGarantiaCredito.TextMatrix(pnRow, 8)) < 0 Then
                MsgBox "El monto de cobertura debe ser mayor a cero", vbInformation, "Aviso"
                Cancel = False
                Exit Sub
            End If
            If CCur(feGarantiaCredito.TextMatrix(pnRow, 8)) > CCur(feGarantiaCredito.TextMatrix(pnRow, 7)) Then
                MsgBox "El monto de cobertura no puede ser mayor a " & feGarantiaCredito.TextMatrix(pnRow, 7), vbInformation, "Aviso"
                Cancel = False
                Exit Sub
            End If
        End If
    End If
End Sub
Private Sub feGarantiaCredito_RowColChange()
    Dim lnMonto As Currency
    If feGarantiaCredito.TextMatrix(1, 0) = "" Then Exit Sub
    
    If feGarantiaCredito.Col = 8 Then
        lnMonto = feGarantiaCredito.SumaRow(15)
        txtMontoCoberturaTC.Text = Format(lnMonto, gsFormatoNumeroView)
    End If
End Sub
Private Sub feGarantiaPersona_DblClick()
    Dim frm As frmGarantia
    If feGarantiaPersona.TextMatrix(1, 0) <> "" Then
        Set frm = New frmGarantia
        frm.Consultar feGarantiaPersona.TextMatrix(feGarantiaPersona.row, 1)
    End If
    Set frm = Nothing
End Sub

Private Sub feGarantiaPersona_GotFocus()
    fbFocoGrilla = True
End Sub
Private Sub feGarantiaPersona_LostFocus()
    fbFocoGrilla = False
End Sub
Private Sub chkGarantiaPersona_Click()

    Dim i As Integer
    Dim lscheck As String

    If feGarantiaPersona.TextMatrix(1, 0) = "" Then
        chkGarantiaPersona.value = 0
        Exit Sub
    End If
    
    If Not fbCheckGrilla Then
        lscheck = IIf(Me.chkGarantiaPersona.value = 1, "1", "0")
        For i = 1 To feGarantiaPersona.Rows - 1
            feGarantiaPersona.TextMatrix(i, 2) = lscheck
        Next
    End If
    checkearGarantias
End Sub
Private Sub feGarantiaPersona_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    Dim lsTpoProdCod As String
    Dim lnTpoBienFuturo As Integer
    Dim lsCtaCod As String
    Dim nCuentaBF As Integer
    Dim j As Integer
    
    Screen.MousePointer = 11
    If Len(Trim(ActxCuenta.NroCuenta)) = 18 Then
        nCuentaBF = 0
        
        Dim oDGarantia As COMNCredito.NCOMGarantia
        Set oDGarantia = New COMNCredito.NCOMGarantia
        
        Dim oRsGarantia As ADODB.Recordset
        Set oRsGarantia = New ADODB.Recordset
        
        Set oRsGarantia = oDGarantia.ObtenerGarantiaColocacionesCobertura(feGarantiaPersona.TextMatrix(pnRow, 1))
        If Not (oRsGarantia.BOF Or oRsGarantia.EOF) Then
            lsTpoProdCod = oRsGarantia!cTpoProdCod
            lnTpoBienFuturo = oRsGarantia!nTpoBienContrato
            lsCtaCod = oRsGarantia!cCtaCod
        End If
        
        For j = 1 To feGarantiaPersona.Rows - 1
            If feGarantiaPersona.TextMatrix(j, 2) = "." Then
                nCuentaBF = nCuentaBF + 1
            End If
        Next
        
        If ((fsTpoProdCod = "802" Or fsTpoProdCod = "806")) Then
            If nCuentaBF > 1 And feGarantiaPersona.TextMatrix(pnRow, 2) = "." Then
                feGarantiaPersona.TextMatrix(pnRow, 2) = ""
                MsgBox "Los créditos con tipo Nuevo Fondo MIVIENDA (FVM) o Techo Propio (TP) solo deben ser coberturados por una sola garantía BIEN FUTURO.", vbExclamation, "Aviso"
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        
        If lnTpoBienFuturo = gTpoBienFuturo Then
            If (Not (fsTpoProdCod = "802" Or fsTpoProdCod = "806")) Then
                feGarantiaPersona.TextMatrix(pnRow, 2) = ""
                MsgBox "La garantía BIEN FUTURO solo cobertura créditos de tipo Nuevo Fondo MIVIENDA (FVM) o Techo Propio (TP).", vbExclamation, "Aviso"
                Screen.MousePointer = 0
                Exit Sub
            End If
            If Len(Trim(lsCtaCod)) = 18 Then
                If feGarantiaPersona.TextMatrix(pnRow, 2) = "." And lsCtaCod <> Trim(ActxCuenta.NroCuenta) Then
                    feGarantiaPersona.TextMatrix(pnRow, 2) = ""
                    MsgBox "La garantía BIEN FUTURO no puede ser seleccionada, ya que viene coberturando al crédito N° " & lsCtaCod & ".", vbExclamation, "Aviso"
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        End If
        If lnTpoBienFuturo <> gTpoBienFuturo Then
             If ((fsTpoProdCod = "802" Or fsTpoProdCod = "806")) Then
                feGarantiaPersona.TextMatrix(pnRow, 2) = ""
                MsgBox "Los créditos con tipo Nuevo Fondo MIVIENDA (FVM) o Techo Propio (TP) deben ser coberturados con garantías BIEN FUTURO.", vbExclamation, "Aviso"
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
    End If
    
    Dim i As Integer
    Dim nCuenta As Integer
    
    On Error GoTo ErrOnCellCheck
    fbCheckGrilla = True
    
    nCuenta = 0
    For i = 1 To feGarantiaPersona.Rows - 1
        If feGarantiaPersona.TextMatrix(i, 2) = "." Then
            nCuenta = nCuenta + 1
        End If
    Next
    
    If (feGarantiaPersona.Rows - 1) = nCuenta Then
        chkGarantiaPersona.value = 1
    Else
        chkGarantiaPersona.value = 0
    End If
    
    fbCheckGrilla = False
    checkearGarantias
    
    Screen.MousePointer = 0
    Exit Sub
ErrOnCellCheck:
    Screen.MousePointer = 0
    fbCheckGrilla = False
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub EditarGarantia(ByVal pnIndex As Integer, _
                                ByVal psNumGarant As String, _
                                ByVal pbSalir As Boolean, _
                                ByVal psNombreTitVerifica As String, _
                                ByRef pbVAL_Migrado As Boolean, _
                                ByRef pbTRA_Migrado As Boolean, _
                                ByRef pnMontoDisp As Double, _
                                ByRef pnMontoAL As Double)
    Dim ofrmGx As New frmGarantia
    Dim oDGx As COMDCredito.DCOMGarantia
    Dim rsGx As ADODB.Recordset
    
    pbSalir = False
    
    If ofrmGx.Editar(psNumGarant) Then
        Set oDGx = New COMDCredito.DCOMGarantia
        Set rsGx = oDGx.RecuperaGarantiaxVerificarActRegCobertura(psNumGarant, fsCtaCod, fsTpoProdCod, fbCliPreferencial)
        Call PintarFilaGarantiaPersona(pnIndex, rsGx)
        
        'Verificamos no haya cambiado de titular->No debería cambiar
        If UCase(Trim(psNombreTitVerifica)) <> UCase(Trim(rsGx!cPersNombreTit)) Then
            pbSalir = True
            MsgBox "Ud. ha cambiado de titular a la Garantía, no se puede continuar.", vbCritical, "Aviso"
            Unload Me
            Exit Sub
        End If
        
        'Actualizar variables usadas en metodo principal
        pbVAL_Migrado = rsGx!bVAL_Migra
        pbTRA_Migrado = rsGx!bTRA_Migra
        'pnMontoDisp = rsGx!nDisponible
        pnMontoDisp = rsGx!nDisponibleGAR 'EJVG20160405
        pnMontoAL = rsGx!nMontoDisponibleAL
        
        RSClose rsGx
    End If
    Set rsGx = Nothing
    Set oDGx = Nothing
    Set ofrmGx = Nothing
End Sub
Private Sub checkearGarantias()
    Dim Index As Integer, i As Integer
    Dim lnRatioPMax As Double, lnRatioNPMax As Double, lnRatio As Double
    Dim lbPreferida As Boolean
    Dim lnOrdenMax As Integer, lnOrden As Integer
    Dim lnMonto As Double, lnTotal As Double
    Dim lnTipoCambio As Double
    Dim lnMontoCobTC As Double, lnMontoCobTemp As Double
    Dim lsTexto As String
    Dim lbVAL_Migrado As Boolean, lbTRA_Migrado As Boolean
    Dim lnMontoDisp As Double, lnMontoAL As Double
    Dim lsNumGarant As String
    Dim lsPersNombreTitGar As String
    Dim lbUnloadXEditarGA As Boolean
    Dim bFirst As Boolean

    On Error GoTo ErrCheck
    Screen.MousePointer = 11
    
    lnTipoCambio = CDbl(txtTipoCambio.Text)
    
    FormateaFlex feGarantiaCredito

    'No mostrar columnas en Próducto RapiFlash
    feGarantiaCredito.ColWidth(6) = 795
    If fsTpoProdCod = "703" Then
        feGarantiaCredito.ColWidth(6) = 0
    End If
    
    If feGarantiaPersona.TextMatrix(1, 0) = "" Then Exit Sub

    lbPreferida = True
    For i = 1 To feGarantiaPersona.Rows - 1
        If feGarantiaPersona.TextMatrix(i, 2) = "." Then
            lsNumGarant = feGarantiaPersona.TextMatrix(i, 1)
            lsPersNombreTitGar = UCase(Trim(feGarantiaPersona.TextMatrix(i, 3)))
            
            If CDbl(feGarantiaPersona.TextMatrix(i, 8)) <= 0# Or CDbl(feGarantiaPersona.TextMatrix(i, 9)) <= 0# Then
                MsgBox "El ratio de cobertura de la Garantía N° " & lsNumGarant & " no esta configurada." & Chr(13) & Chr(13) & "Comuniquese con la Gerencia de Riesgos." & _
                        Chr(13) & Chr(13) & "- Tipo Producto: " & txtProductoDesc.Text & _
                        Chr(13) & "- Tipo Bien: " & feGarantiaPersona.TextMatrix(i, 5) & _
                        Chr(13) & Chr(13) & "Luego que lo hayan configurado, volver a cargar el crédito.", vbInformation, "Aviso"
                feGarantiaPersona.TextMatrix(i, 2) = ""
                chkGarantiaPersona.value = 0 'Quitamos los checks actuales
                chkGarantiaPersona_Click
                Exit Sub
            End If
            
            'Validamos se hayan actualizado las valorizaciones y tramites legales, que no sean las migradas por el sistema
            lbVAL_Migrado = IIf(CInt(feGarantiaPersona.TextMatrix(i, 18)) = 1, True, False)
            lbTRA_Migrado = IIf(CInt(feGarantiaPersona.TextMatrix(i, 19)) = 1, True, False)
            
            lsTexto = ""
            bFirst = True
            While (lbVAL_Migrado Or lbTRA_Migrado)
                If Not bFirst Then
                    feGarantiaPersona.TextMatrix(i, 2) = ""
                    chkGarantiaPersona.value = 0 'Quitamos los checks actuales
                    chkGarantiaPersona_Click
                    Exit Sub
                End If
                bFirst = False
                
                lsTexto = "La Garantía N° " & lsNumGarant & " necesita ser actualizada en:" & Chr(13) & Chr(13)
                
                If lbVAL_Migrado Then
                    lsTexto = lsTexto & "- VALORIZACIÓN -> La última valorización fue una migración del Sistema." & Chr(13)
                End If
                If lbTRA_Migrado Then
                    lsTexto = lsTexto & "- TRÁMITE LEGAL -> El último Trámite Legal fue una migración del Sistema." & Chr(13)
                End If

                lsTexto = lsTexto & Chr(13) & "¿Desea realizar la actualización de la información ahora?"

                If MsgBox(lsTexto, vbInformation + vbYesNo, "Aviso") = vbYes Then
                    Call EditarGarantia(i, lsNumGarant, lbUnloadXEditarGA, lsPersNombreTitGar, lbVAL_Migrado, lbTRA_Migrado, lnMontoDisp, lnMontoAL)
                    If lbUnloadXEditarGA Then Exit Sub
                End If
            Wend
            
            lnMontoAL = CDbl(feGarantiaPersona.TextMatrix(i, 17))
            lnMontoDisp = CDbl(feGarantiaPersona.TextMatrix(i, 21))
            
            lsTexto = ""
            bFirst = True
            'Para las Garantías AutoLiquidables el disponible a la fecha debe ser igual al de la última valorización, sino no puede continuar. De manera que se guarde el historico de las valorizaciones
            If fsTpoProdCod = "703" Or lnMontoAL > 0# Then
                While (lnMontoAL <> lnMontoDisp)
                    If Not bFirst Then
                        feGarantiaPersona.TextMatrix(i, 2) = ""
                        chkGarantiaPersona.value = 0 'Quitamos los checks actuales
                        chkGarantiaPersona_Click
                        Exit Sub
                    End If
                    bFirst = False
                    
                    lsTexto = "¡El disponible de la Garantía N° " & lsNumGarant & " no es igual al de la última valuación!." & Chr(13) & Chr(13) & _
                            "- Monto disponible Garantía:" & Space(16) & Format(lnMontoDisp, gsFormatoNumeroView) & Chr(13) & _
                            "- Monto disponible última Valuación: " & Space(3) & Format(lnMontoAL, gsFormatoNumeroView) & Chr(13) & Chr(13) & _
                            "Edite la garantía y adicione una nueva Valorización." & Chr(13) & Chr(13) & "¿Desea realizar la actualización de la información ahora?"
                    
                    If MsgBox(lsTexto, vbInformation + vbYesNo, "Aviso") = vbYes Then
                        Call EditarGarantia(i, lsNumGarant, lbUnloadXEditarGA, lsPersNombreTitGar, lbVAL_Migrado, lbTRA_Migrado, lnMontoDisp, lnMontoAL)
                        If lbUnloadXEditarGA Then Exit Sub
                    End If
                Wend
            End If
            
            feGarantiaCredito.AdicionaFila
            Index = feGarantiaCredito.row
            
            If CDbl(feGarantiaPersona.TextMatrix(i, 8)) > lnRatioPMax Then
                lnRatioPMax = CDbl(feGarantiaPersona.TextMatrix(i, 8))
            End If
            If CDbl(feGarantiaPersona.TextMatrix(i, 9)) > lnRatioNPMax Then
                lnRatioNPMax = CDbl(feGarantiaPersona.TextMatrix(i, 9))
            End If
            If CInt(feGarantiaPersona.TextMatrix(i, 13)) > lnOrdenMax Then
                lnOrdenMax = CInt(feGarantiaPersona.TextMatrix(i, 13))
            End If
            
            If lbPreferida Then
                If feGarantiaPersona.TextMatrix(i, 12) = "0" Then 'No es Preferida
                    lbPreferida = False
                End If
            End If
            
            feGarantiaCredito.TextMatrix(Index, 1) = lsNumGarant 'ID Garantia
            feGarantiaCredito.TextMatrix(Index, 2) = feGarantiaPersona.TextMatrix(i, 3) 'Titular
            feGarantiaCredito.TextMatrix(Index, 3) = lsNumGarant 'ID Garantia Mostrar
            feGarantiaCredito.TextMatrix(Index, 4) = feGarantiaPersona.TextMatrix(i, 5) 'Bien
            feGarantiaCredito.TextMatrix(Index, 5) = feGarantiaPersona.TextMatrix(i, 7) 'Saldo Cob
            feGarantiaCredito.TextMatrix(Index, 6) = "0.00" 'Ratio
            feGarantiaCredito.TextMatrix(Index, 7) = "0.00" 'Max Cob
            feGarantiaCredito.TextMatrix(Index, 8) = "0.00" 'Monto Cob
            feGarantiaCredito.TextMatrix(Index, 9) = feGarantiaPersona.TextMatrix(i, 13) 'Orden
            feGarantiaCredito.TextMatrix(Index, 10) = feGarantiaPersona.TextMatrix(i, 14) 'Moneda
            If feGarantiaPersona.TextMatrix(i, 14) = Moneda.gMonedaNacional Then
                feGarantiaCredito.BackColorRow &HC0FFFF 'vbYellow
            ElseIf feGarantiaPersona.TextMatrix(i, 14) = Moneda.gMonedaExtranjera Then
                feGarantiaCredito.BackColorRow &HC0FFC0 'vbGreen
            End If
            feGarantiaCredito.TextMatrix(Index, 11) = feGarantiaPersona.TextMatrix(i, 15) 'FechaValorizacion
            feGarantiaCredito.TextMatrix(Index, 12) = feGarantiaPersona.TextMatrix(i, 16) 'FechaTramiteLegal
            feGarantiaCredito.TextMatrix(Index, 13) = i 'IndexGarantiaPersona
            feGarantiaCredito.TextMatrix(Index, 14) = "0.00" 'Monto Cob Temporal
            feGarantiaCredito.TextMatrix(Index, 15) = "0.00" 'Monto Cob Temporal
        End If
    Next
    feGarantiaCredito.row = 1
    feGarantiaCredito.TopRow = 1
    
    If feGarantiaCredito.TextMatrix(1, 0) <> "" Then
        'Llenamos Ratio(En caso que tenga exoneración se tomará de lo establecido)
        lnRatio = IIf(lbPreferida, lnRatioPMax, lnRatioNPMax)
        'lnRatio = IIf(fnExoneraID > 0 And fnExoneraAprueba = 1, IIf(fsTpoProdCod = "703", 1 / IIf(fnExoneraTasa = 0, 1, fnExoneraTasa), fnExoneraTasa), IIf(fsTpoProdCod = "703", 1 / IIf(lnRatio = 0, 1, lnRatio), lnRatio)) 'RECO20160713 ERS002-2016
        lnRatio = IIf(fnExoneraTasa > 0, fnExoneraTasa, IIf(fsTpoProdCod = "703", 1 / IIf(lnRatio = 0, 1, lnRatio), lnRatio)) 'RECO20160713 ERS002-2016
        'Llenamos Max. Cobertura
        For Index = 1 To feGarantiaCredito.Rows - 1
            feGarantiaCredito.TextMatrix(Index, 6) = Format(lnRatio, "0.00##############") 'Ratio->No poner formato para mantener cantidad de decimales
            feGarantiaCredito.TextMatrix(Index, 7) = Format(Round(CDbl(feGarantiaCredito.TextMatrix(Index, 5)) / lnRatio, 2), gsFormatoNumeroView) 'Max Cob
        Next
        'Llenamos Monto Cobertura por defecto
        lnMonto = CDbl(txtMonto.Text)
        For lnOrden = lnOrdenMax To 1 Step -1
            For Index = 1 To feGarantiaCredito.Rows - 1
                If CInt(feGarantiaCredito.TextMatrix(Index, 9)) = lnOrden Then
                    If fnMoneda <> CInt(feGarantiaCredito.TextMatrix(Index, 10)) Then 'Comparamos moneda del crédito con la garantia
                        If fnMoneda = gMonedaNacional Then 'Garantia Dolares -> Crédito Soles
                            lnMontoCobTC = Round(CDbl(feGarantiaCredito.TextMatrix(Index, 7)) * lnTipoCambio, 2)
                            If lnMonto >= lnMontoCobTC Then
                                lnMonto = lnMonto - lnMontoCobTC
                                feGarantiaCredito.TextMatrix(Index, 8) = feGarantiaCredito.TextMatrix(Index, 7) 'Monto Cob
                                feGarantiaCredito.TextMatrix(Index, 14) = feGarantiaCredito.TextMatrix(Index, 7) 'Monto Cob Temporal
                                feGarantiaCredito.TextMatrix(Index, 15) = Format(lnMontoCobTC, gsFormatoNumeroView) 'Monto Cob Tipo Cambio
                                'lnTotal = lnTotal + lnMontoCobTC
                            ElseIf lnMonto < lnMontoCobTC And lnMonto > 0 Then
                                lnMontoCobTemp = Round(lnMonto / lnTipoCambio, 2)
                                feGarantiaCredito.TextMatrix(Index, 8) = Format(lnMontoCobTemp, gsFormatoNumeroView) 'Monto Cob
                                feGarantiaCredito.TextMatrix(Index, 14) = Format(lnMontoCobTemp, gsFormatoNumeroView) 'Monto Cob Temporal
                                feGarantiaCredito.TextMatrix(Index, 15) = Format(Round(lnMontoCobTemp * lnTipoCambio, 2), gsFormatoNumeroView) 'Monto Cob Tipo Cambio
                                'lnTotal = lnTotal + Round(lnMontoCobTemp * lnTipoCambio, 2)
                                lnMonto = 0 'lnMonto - lnMontoCobTC
                            End If
                        ElseIf fnMoneda = gMonedaExtranjera Then 'Garantia Soles -> Crédito Dolares
                            lnMontoCobTC = Round(CDbl(feGarantiaCredito.TextMatrix(Index, 7)) / lnTipoCambio, 2)
                            If lnMonto >= lnMontoCobTC Then
                                lnMonto = lnMonto - lnMontoCobTC
                                feGarantiaCredito.TextMatrix(Index, 8) = feGarantiaCredito.TextMatrix(Index, 7) 'Monto Cob
                                feGarantiaCredito.TextMatrix(Index, 14) = feGarantiaCredito.TextMatrix(Index, 7) 'Monto Cob Temporal
                                feGarantiaCredito.TextMatrix(Index, 15) = Format(lnMontoCobTC, gsFormatoNumeroView) 'Monto Cob Tipo Cambio
                                'lnTotal = lnTotal + lnMontoCobTC
                            ElseIf lnMonto < lnMontoCobTC And lnMonto > 0 Then
                                lnMontoCobTemp = Round(lnMonto * lnTipoCambio, 2)
                                feGarantiaCredito.TextMatrix(Index, 8) = Format(lnMontoCobTemp, gsFormatoNumeroView) 'Monto Cob
                                feGarantiaCredito.TextMatrix(Index, 14) = Format(lnMontoCobTemp, gsFormatoNumeroView) 'Monto Cob Temporal
                                feGarantiaCredito.TextMatrix(Index, 15) = Format(Round(lnMontoCobTemp / lnTipoCambio, 2), gsFormatoNumeroView) 'Monto Cob Tipo Cambio
                                'lnTotal = lnTotal + Round(lnMontoCobTemp / lnTipoCambio, 2)
                                lnMonto = 0 'lnMonto - lnMontoCobTC
                            End If
                        End If
                    Else
                        lnMontoCobTC = CDbl(feGarantiaCredito.TextMatrix(Index, 7))
                        If lnMonto >= lnMontoCobTC Then
                            lnMonto = lnMonto - lnMontoCobTC
                            feGarantiaCredito.TextMatrix(Index, 8) = feGarantiaCredito.TextMatrix(Index, 7) 'Monto Cob
                            feGarantiaCredito.TextMatrix(Index, 14) = feGarantiaCredito.TextMatrix(Index, 7) 'Monto Cob Temporal
                            feGarantiaCredito.TextMatrix(Index, 15) = Format(lnMontoCobTC, gsFormatoNumeroView) 'Monto Cob Tipo Cambio
                            'lnTotal = lnTotal + lnMontoCobTC
                        ElseIf lnMonto < lnMontoCobTC And lnMonto > 0 Then
                            lnMontoCobTemp = lnMonto 'Round(lnMonto * lnTipoCambio, 2)
                            feGarantiaCredito.TextMatrix(Index, 8) = Format(lnMontoCobTemp, gsFormatoNumeroView) 'Monto Cob
                            feGarantiaCredito.TextMatrix(Index, 14) = Format(lnMontoCobTemp, gsFormatoNumeroView) 'Monto Cob Temporal
                            feGarantiaCredito.TextMatrix(Index, 15) = Format(Round(lnMontoCobTemp, 2), gsFormatoNumeroView) 'Monto Cob Tipo Cambio
                            'lnTotal = lnTotal + Round(lnMontoCobTemp / lnTipoCambio, 2)
                            lnMonto = 0 'lnMonto - lnMontoCobTC
                        End If
                    End If
                End If
            Next
        Next
        If EnfocaControl(feGarantiaCredito) Then
            feGarantiaCredito.Col = 8
        End If
    End If
    
    txtMontoCoberturaTC.Text = Format(feGarantiaCredito.SumaRow(15), gsFormatoNumeroView)
    'DeterminaAccionExoneracion 'RECO20160628 ERS002-2016
    
    Screen.MousePointer = 0
    Exit Sub
ErrCheck:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
    'Deschekeamos todos los registros
    If i > 0 Then
        feGarantiaPersona.TextMatrix(i, 2) = ""
    End If
    chkGarantiaPersona.value = 0 'Quitamos los checks actuales
    chkGarantiaPersona_Click
End Sub
Private Sub feGarantiaPersona_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim bCancel As Boolean
    bCancel = EsCeldaFlexEditable(feGarantiaPersona, pnCol)
    If Not bCancel Then
        Cancel = bCancel
        Exit Sub
    End If
End Sub
Private Sub fePersona_GotFocus()
    fbFocoGrilla = True
End Sub
Private Sub fePersona_LostFocus()
    fbFocoGrilla = False
End Sub
Private Sub Form_Activate()
    If fbSalir Then
        Unload Me
        Screen.MousePointer = 0
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If fbFocoGrilla Then
        If KeyCode = 86 And Shift = 2 Then
            KeyCode = 10
        End If
    End If
End Sub
Private Sub Form_Load()
    fbAceptar = False
    Limpiar
    cargarControles
    
    If fnInicio = InicioGravamenxMenu Then
        cmdCancelar.Enabled = True
    ElseIf fnInicio = InicioGravamenxSolicitud Then
        ActxCuenta.NroCuenta = fsCtaCod
        HabilitarBusqueda False
        
        If Not CargaDatos(ActxCuenta.NroCuenta) Then
            MsgBox "Credito no puede ser seleccionado o no existe", vbInformation, "Aviso"
            fbSalir = True
            Exit Sub
        Else
            If (fnExoneraID > 0 And fnExoneraAprueba = -1) Then
                fbSalir = True
                Exit Sub
            Else
                fbSalir = False
            End If
        End If
        
        cmdSolicitud.Enabled = True
        cmdSugerencia.Enabled = fbDataGarantiaCredito
    ElseIf fnInicio = InicioGravamenxAjuste Then
        ActxCuenta.NroCuenta = fsCtaCod
        HabilitarBusqueda False
        
        If Not CargaDatos(ActxCuenta.NroCuenta) Then
            MsgBox "Credito No Existe", vbInformation, "Aviso"
            fbSalir = True
            Exit Sub
        Else
            If (fnExoneraID > 0 And fnExoneraAprueba = -1) Then
                fbSalir = True
                Exit Sub
            End If
        End If
    End If
    
    If fnTipoCamb = 0 Then
        MsgBox "No existe tipo de cambio a la fecha", vbInformation, "Aviso"
        fbSalir = True
        Exit Sub
    End If
    
    gsOpeCod = gCredRegistrarGravamen 'Log Grabar Cobertura
End Sub
Private Sub cargarControles()
    Dim oTipCambio As New COMDConstSistema.NCOMTipoCambio
    fnTipoCamb = oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoDia)
    txtTipoCambio.Text = Format(fnTipoCamb, gsFormatoNumeroView)
    Set oTipCambio = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    fbSalir = False
End Sub

