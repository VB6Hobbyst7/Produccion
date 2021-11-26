VERSION 5.00
Begin VB.Form frmCapPersParam 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9225
   Icon            =   "frmCapPersParam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
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
      Height          =   735
      Left            =   5760
      TabIndex        =   9
      Top             =   90
      Visible         =   0   'False
      Width           =   3300
      Begin VB.ComboBox cboPrograma 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   270
         Width           =   3030
      End
   End
   Begin VB.Frame fraOrdenPago 
      Caption         =   "Orden Pago"
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
      Height          =   735
      Left            =   4050
      TabIndex        =   7
      Top             =   90
      Width           =   1680
      Begin VB.CheckBox chkOrdenPago 
         Caption         =   "&Orden Pago"
         Height          =   240
         Left            =   405
         TabIndex        =   8
         Top             =   315
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   4365
      Width           =   1230
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   4860
      TabIndex        =   3
      Top             =   4365
      Width           =   1185
   End
   Begin VB.Frame fraParametro 
      Caption         =   "Parámetros"
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
      Height          =   3390
      Left            =   90
      TabIndex        =   6
      Top             =   900
      Width           =   8970
      Begin SICMACT.FlexEdit grdParam 
         Height          =   2985
         Left            =   90
         TabIndex        =   2
         Top             =   225
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   5265
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Personería-Mon Min Apert-Saldo Minimo-Mon Min Dep-Mon Min Ret"
         EncabezadosAnchos=   "350-2800-1300-1300-1300-1300"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-4-5"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-R-C"
         FormatosEdit    =   "0-0-2-2-2-2"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraMoneda 
      Caption         =   "Moneda"
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
      Height          =   735
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   3885
      Begin VB.OptionButton optMoneda 
         Caption         =   "Moneda &Extranjera"
         Height          =   285
         Index           =   1
         Left            =   2070
         TabIndex        =   1
         Top             =   270
         Width           =   1680
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "Moneda &Nacional"
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   0
         Top             =   270
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmCapPersParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nProducto As Producto
Dim nMoneda As Moneda
Dim bOrdPag As Boolean
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by


Private Sub ObtieneDatosParametros()
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
grdParam.Clear
grdParam.Rows = 2
grdParam.FormaCabecera
Set grdParam.Recordset = oCap.GetCapPersoneriaParam(nProducto, nMoneda, bOrdPag)
Set oCap = Nothing
grdParam.FormateaColumnas
End Sub

Public Sub Inicia(ByVal nProd As Producto)
nProducto = nProd
fraOrdenPago.Visible = False
Select Case nProducto
    Case gCapAhorros
        Me.Caption = "Captaciones - Parámetros por Personería - Ahorros"
        fraOrdenPago.Visible = True
        'By Capi 20012009
        Set objPista = New COMManejador.Pista
        gsOpeCod = gAhoParamPeroneria
        'End By


    Case gCapPlazoFijo
        Me.Caption = "Captaciones - Parámetros por Personería - Plazo Fijo"
         'By Capi 20012009
        Set objPista = New COMManejador.Pista
        gsOpeCod = gPFParamPeroneria
        'End By
    Case gCapCTS
        Me.Caption = "Captaciones - Parámetros por Personería - CTS"
         'By Capi 20012009
        Set objPista = New COMManejador.Pista
        gsOpeCod = gCTSParamPeroneria
        'End By
End Select
Me.Show 1
End Sub

Private Sub chkOrdenPago_Click()
bOrdPag = IIf(chkOrdenPago.value = 1, True, False)
ObtieneDatosParametros
End Sub

Private Sub cmdGrabar_Click()
Dim oCont As COMNContabilidad.NCOMContFunciones
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim sMovNro As String
Dim rsPar As ADODB.Recordset

If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Set oCont = New COMNContabilidad.NCOMContFunciones
        sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oCont = Nothing
    Set rsPar = New ADODB.Recordset
    Set rsPar = grdParam.GetRsNew()
    
    Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion
        oCap.ActualizaCapPersoneriaParam rsPar, nProducto, nMoneda, bOrdPag, sMovNro
    'By Capi 21012009
     objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Parametros Personeria"
    'End by
            
   
        
    Set oCap = Nothing
    Set rsPar = Nothing
    
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    optMoneda(0).value = True
    chkOrdenPago.value = 0
    bOrdPag = False
End Sub

Private Sub optMoneda_Click(Index As Integer)
    If Index = 0 Then
        nMoneda = gMonedaNacional
    ElseIf Index = 1 Then
        nMoneda = gMonedaExtranjera
    End If
    ObtieneDatosParametros
End Sub
