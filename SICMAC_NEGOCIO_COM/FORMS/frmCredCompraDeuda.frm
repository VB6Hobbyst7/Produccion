VERSION 5.00
Begin VB.Form frmCredCompraDeuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compra de deuda a otras Instituciones Financieras"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "frmCredCompraDeuda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Datos de Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3470
      Left            =   40
      TabIndex        =   0
      Top             =   40
      Width           =   8700
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   2475
         TabIndex        =   5
         ToolTipText     =   "Eliminar"
         Top             =   3020
         Width           =   1000
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   345
         Left            =   7560
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   3020
         Width           =   1000
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Nuevo"
         Height          =   345
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Nuevo"
         Top             =   3020
         Width           =   1000
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Editar"
         Height          =   345
         Left            =   1160
         TabIndex        =   1
         ToolTipText     =   "Editar"
         Top             =   3020
         Width           =   1000
      End
      Begin SICMACT.FlexEdit feDeudaComprar 
         Height          =   2730
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   4815
         Cols0           =   8
         HighLight       =   2
         EncabezadosNombres=   "N°-Institución Financiera-Nro. Crédito-Moneda-N°Cuotas-Saldo Comprar-Monto Cuota-Aux"
         EncabezadosAnchos=   "400-2200-1500-800-900-1300-1200-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-C-R-R-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         CantEntero      =   12
         TextArray0      =   "N°"
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmCredCompraDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'** Nombre : frmCredCompraDeuda
'** Descripción : Para listar las deudas que se comprará a otras IFIs creado segun TI-ERS002-2016
'** Creación : EJVG, 20160129 04:30:00 PM
'************************************************************************************************
Option Explicit

Dim fvListaCompraDeuda() As TCompraDeuda
Dim fnTpoProducto As Integer '**ARLO20180317 ERS 070 - 2017 ANEXO 02

Private Sub Form_Load()
    Dim i As Integer
    
    cmdQuitar.Enabled = False
    cmdModificar.Enabled = False
    
    If UBound(fvListaCompraDeuda) > 0 Then
        For i = 1 To UBound(fvListaCompraDeuda)
            Call AdicionaFila(fvListaCompraDeuda(i))
        Next
        cmdQuitar.Enabled = True
        cmdModificar.Enabled = True
    End If
End Sub
Public Sub Inicio(ByRef pvListaCompraDeuda() As TCompraDeuda, ByVal psnTpoProducto As Integer) '**ARLO20180317 ADD psCtpoProducto
    fvListaCompraDeuda = pvListaCompraDeuda
    fnTpoProducto = psnTpoProducto '**ARLO20180317 ADD psCtpoProducto
    Show 1
    pvListaCompraDeuda = fvListaCompraDeuda
End Sub
Private Sub cmdAgregar_Click()
    Dim frm As New frmCredCompraDeudaDet
    Dim lvCompraDeuda As TCompraDeuda
    Dim lvTemp() As TCompraDeuda
    Dim bOK As Boolean
    Dim Index As Integer
    
    lvTemp = fvListaCompraDeuda 'Temporal para no modificar el actual array
    
    bOK = frm.Registrar(lvCompraDeuda, lvTemp, fnTpoProducto) '**ARLO20180317 ADD fnTpoProducto
    If bOK Then
        Index = UBound(fvListaCompraDeuda) + 1
        ReDim Preserve fvListaCompraDeuda(Index)
        fvListaCompraDeuda(Index) = lvCompraDeuda
        
        AdicionaFila lvCompraDeuda
        
        cmdQuitar.Enabled = True
        cmdModificar.Enabled = True
    End If
    Set frm = Nothing
End Sub
Private Sub cmdModificar_Click()
    Dim frm As frmCredCompraDeudaDet
    Dim lvCompraDeuda As TCompraDeuda
    Dim lvTemp() As TCompraDeuda
    Dim bOK As Boolean
    Dim Index As Integer
    
    If feDeudaComprar.TextMatrix(1, 0) = "" Then Exit Sub
    
    Index = feDeudaComprar.row
    lvTemp = fvListaCompraDeuda 'Temporal para no modificar el actual array
    
    lvCompraDeuda = fvListaCompraDeuda(Index)
    Set frm = New frmCredCompraDeudaDet
    
    bOK = frm.Modificar(lvCompraDeuda, Index, lvTemp, fnTpoProducto) '**ARLO20180317 ADD fnTpoProducto)
    If bOK Then
        fvListaCompraDeuda(Index) = lvCompraDeuda
        ModificaFila Index, lvCompraDeuda
    End If
    Set frm = Nothing
End Sub
Private Sub cmdQuitar_Click()
    Dim lvListaTemp() As TCompraDeuda
    Dim Index As Integer
    Dim i As Integer
    Dim j As Integer
        
    If feDeudaComprar.TextMatrix(1, 0) = "" Then Exit Sub
    
    Index = feDeudaComprar.row
    
    If MsgBox("Se va a eliminar la compra de deuda." & Chr(13) & Chr(13) & "Institución Financiera: " & fvListaCompraDeuda(Index).sIFINombre & Chr(13) & "Nro. Crédito : " & fvListaCompraDeuda(Index).sNroCredito & Chr(13) & Chr(13) & "¿Desea continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    ReDim lvListaTemp(0)
    
    For i = 1 To UBound(fvListaCompraDeuda)
        If i <> Index Then
            j = UBound(lvListaTemp) + 1
            ReDim Preserve lvListaTemp(j)
            lvListaTemp(j) = fvListaCompraDeuda(i)
        End If
    Next
    
    fvListaCompraDeuda = lvListaTemp
    feDeudaComprar.EliminaFila Index
    
    If UBound(fvListaCompraDeuda) = 0 Then
        cmdQuitar.Enabled = False
        cmdModificar.Enabled = False
    End If
End Sub
Private Sub feDeudaComprar_RowColChange()
    If feDeudaComprar.TextMatrix(1, 0) = "" Then
        cmdModificar.Enabled = False
    Else
        cmdModificar.Enabled = True
    End If
End Sub
Private Sub feDeudaComprar_DblClick()
    If feDeudaComprar.TextMatrix(1, 0) = "" Then Exit Sub
    cmdModificar_Click
End Sub
Private Sub AdicionaFila(ByRef pvCompraDeuda As TCompraDeuda)
    Dim i As Integer
    
    feDeudaComprar.AdicionaFila
    i = feDeudaComprar.row
    feDeudaComprar.TextMatrix(i, 1) = pvCompraDeuda.sIFINombre
    feDeudaComprar.TextMatrix(i, 2) = pvCompraDeuda.sNroCredito
    feDeudaComprar.TextMatrix(i, 3) = IIf(pvCompraDeuda.nMoneda = gMonedaNacional, "SOLES", "DOLARES")
    feDeudaComprar.TextMatrix(i, 4) = pvCompraDeuda.nNroCuotasPactadas
    feDeudaComprar.TextMatrix(i, 5) = Format(pvCompraDeuda.nSaldoComprar, "#,##0.00")
    feDeudaComprar.TextMatrix(i, 6) = Format(pvCompraDeuda.nMontoCuota, "#,##0.00")
    
    feDeudaComprar.BackColorRow vbWhite
    
    If pvCompraDeuda.nMoneda = gMonedaExtranjera Then
        feDeudaComprar.BackColorRow vbGreen
    End If
End Sub
Private Sub ModificaFila(ByVal pnIndex As Integer, ByRef pvCompraDeuda As TCompraDeuda)
    
    feDeudaComprar.TextMatrix(pnIndex, 1) = pvCompraDeuda.sIFINombre
    feDeudaComprar.TextMatrix(pnIndex, 2) = pvCompraDeuda.sNroCredito
    feDeudaComprar.TextMatrix(pnIndex, 3) = IIf(pvCompraDeuda.nMoneda = gMonedaNacional, "SOLES", "DOLARES")
    feDeudaComprar.TextMatrix(pnIndex, 4) = pvCompraDeuda.nNroCuotasPactadas
    feDeudaComprar.TextMatrix(pnIndex, 5) = Format(pvCompraDeuda.nSaldoComprar, "#,##0.00")
    feDeudaComprar.TextMatrix(pnIndex, 6) = Format(pvCompraDeuda.nMontoCuota, "#,##0.00")
    
    feDeudaComprar.BackColorRow vbWhite
    
    If pvCompraDeuda.nMoneda = gMonedaExtranjera Then
        feDeudaComprar.BackColorRow vbGreen
    End If
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
