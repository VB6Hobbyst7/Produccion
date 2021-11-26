VERSION 5.00
Begin VB.Form frmLogActaConformidadBusca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Órdenes de Compra/Servicio"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   Icon            =   "frmLogActaConformidadBusca.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   8775
      TabIndex        =   1
      Top             =   3435
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   9810
      TabIndex        =   2
      Top             =   3435
      Width           =   1050
   End
   Begin Sicmact.FlexEdit feOrden 
      Height          =   3300
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   5821
      Cols0           =   9
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-nMovNro-Número-N° OC/S Di/Pr-Fecha-Proveedor-Moneda-Importe-Observaciones"
      EncabezadosAnchos=   "400-0-1400-1500-1000-2300-1000-1200-7000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-C-L-C-R-L"
      FormatosEdit    =   "0-0-0-0-0-0-2-2-0"
      TextArray0      =   "#"
      SelectionMode   =   1
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   7
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin Sicmact.FlexEdit feContrato 
      Height          =   3300
      Left            =   75
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   5821
      Cols0           =   8
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-N° Contrato-Proveedor-Moneda-Monto-Desde-Hasta-N° Cuotas"
      EncabezadosAnchos=   "400-1200-2300-1000-1000-900-900-900"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-2-0-0-0"
      TextArray0      =   "#"
      SelectionMode   =   1
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   7
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmLogActaConformidadBusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'** Nombre : frmLogActaConformidad
'** Descripción : Registro de Acta de Conformidad creado segun ERS062-2013
'** Creación : EJVG, 20131009 09:00:00 AM
'***************************************************************************
Option Explicit
Dim fnTpoBusqueda As Integer
Dim fsDato As String
Dim fsDocNro As String
Dim fsDocNombre As String
Dim fRs As New ADODB.Recordset

Private Sub Form_Load()
    Dim row As Long
    Screen.MousePointer = 11
    Select Case fnTpoBusqueda
        Case LogTipoDocOrigenActaConformidad.OrdenCompra, LogTipoDocOrigenActaConformidad.OrdenServicio
            If fnTpoBusqueda = LogTipoDocOrigenActaConformidad.OrdenCompra Then
                Caption = "BUSQUEDA DE ORDENES DE COMPRA"
            Else
                Caption = "BUSQUEDA DE ORDENES DE SERVICIO"
            End If
            feOrden.Visible = True
            LimpiaFlex feOrden
            Do While Not fRs.EOF
                feOrden.AdicionaFila
                row = feOrden.row
                feOrden.TextMatrix(row, 1) = fRs!nMovNro
                feOrden.TextMatrix(row, 2) = fRs!cDocNro
                feOrden.TextMatrix(row, 3) = fRs!cDocNroODirPro
                feOrden.TextMatrix(row, 4) = Format(fRs!dDocFecha, "dd/mm/yyyy")
                feOrden.TextMatrix(row, 5) = fRs!cProveedorNombre
                feOrden.TextMatrix(row, 6) = fRs!cMoneda
                feOrden.TextMatrix(row, 7) = Format(fRs!nImporte, gsFormatoNumeroView)
                feOrden.TextMatrix(row, 8) = fRs!cMovDesc
                fRs.MoveNext
            Loop
            If fRs.RecordCount > 0 Then
                feOrden.TabIndex = 0
                cmdAceptar.Default = True
            Else
                cmdAceptar.Default = False
            End If
            SendKeys "{Right}"
        Case LogTipoDocOrigenActaConformidad.ContratoCompra, LogTipoDocOrigenActaConformidad.ContratoServicio
            If fnTpoBusqueda = LogTipoDocOrigenActaConformidad.ContratoCompra Then
                Caption = "BUSQUEDA DE CONTRATOS DE COMPRA"
            Else
                Caption = "BUSQUEDA DE CONTRATOS DE SERVICIO"
            End If
            feContrato.Visible = True
            LimpiaFlex feContrato
            Do While Not fRs.EOF
                feContrato.AdicionaFila
                row = feContrato.row
                feContrato.TextMatrix(row, 1) = fRs!cNContrato
                feContrato.TextMatrix(row, 2) = fRs!cPersNombre
                feContrato.TextMatrix(row, 3) = fRs!cMoneda
                feContrato.TextMatrix(row, 4) = Format(fRs!nMonto, gsFormatoNumeroView)
                feContrato.TextMatrix(row, 5) = Format(fRs!dFechaIni, gsFormatoFechaView)
                feContrato.TextMatrix(row, 6) = Format(fRs!dFechaFin, gsFormatoFechaView)
                fRs.MoveNext
            Loop
            If fRs.RecordCount > 0 Then
                feContrato.TabIndex = 0
                cmdAceptar.Default = True
            Else
                cmdAceptar.Default = False
            End If
            SendKeys "{Right}"
    End Select
    Screen.MousePointer = 0
End Sub
Public Sub Inicio(ByVal pnTpoBusqueda As Integer, ByRef psDato As String, ByRef psDocumentoCod As String, ByRef psDocumentoNombre As String, ByVal pRs As ADODB.Recordset)
    fnTpoBusqueda = pnTpoBusqueda
    Set fRs = pRs
    Show 1
    psDato = fsDato
    psDocumentoCod = fsDocNro
    psDocumentoNombre = fsDocNombre
End Sub
Private Sub feOrden_DblClick()
    If feOrden.row > 0 Then
        cmdAceptar_Click
    End If
End Sub
Private Sub feContrato_DblClick()
    If feContrato.row > 0 Then
        cmdAceptar_Click
    End If
End Sub
Private Sub cmdAceptar_Click()
    Dim lsMensaje As String
    Dim row As Long
    If fnTpoBusqueda = LogTipoDocOrigenActaConformidad.OrdenCompra Or fnTpoBusqueda = LogTipoDocOrigenActaConformidad.OrdenServicio Then
        row = feOrden.row
        fsDato = CStr(feOrden.TextMatrix(row, 1))
        fsDocNro = feOrden.TextMatrix(row, 2)
        fsDocNombre = feOrden.TextMatrix(row, 5)
    ElseIf fnTpoBusqueda = LogTipoDocOrigenActaConformidad.ContratoCompra Or fnTpoBusqueda = LogTipoDocOrigenActaConformidad.ContratoServicio Then
        row = feContrato.row
        fsDato = feContrato.TextMatrix(row, 1)
        fsDocNro = feContrato.TextMatrix(row, 1)
        fsDocNombre = feContrato.TextMatrix(row, 2)
    End If
    If fsDato = "" Then
        If fnTpoBusqueda = LogTipoDocOrigenActaConformidad.OrdenCompra Then
            lsMensaje = "Ud. debe seleccionar primero la Orden de Compra"
        ElseIf fnTpoBusqueda = LogTipoDocOrigenActaConformidad.OrdenServicio Then
            lsMensaje = "Ud. debe seleccionar primero la Orden de Servicio"
        ElseIf fnTpoBusqueda = LogTipoDocOrigenActaConformidad.ContratoCompra Then
            lsMensaje = "Ud. debe seleccionar primero el Contrato de Compra"
        ElseIf fnTpoBusqueda = LogTipoDocOrigenActaConformidad.ContratoServicio Then
            lsMensaje = "Ud. debe seleccionar primero el Contrato de Servicio"
        End If
        
        MsgBox lsMensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
