VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAdeudPagLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adeudados: Pago en Lote"
   ClientHeight    =   5880
   ClientLeft      =   585
   ClientTop       =   2040
   ClientWidth     =   10260
   Icon            =   "frmAdeudPagLote.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Movimiento"
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
      Left            =   90
      TabIndex        =   14
      Top             =   -30
      Width           =   3960
      Begin VB.TextBox txtOpeCod 
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
         Left            =   930
         TabIndex        =   15
         Top             =   195
         Width           =   900
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   2550
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   195
         Width           =   1275
         _ExtentX        =   2249
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
         Height          =   195
         Left            =   1995
         TabIndex        =   17
         Top             =   225
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Operación"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   225
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "&Procesar"
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
      Left            =   8670
      TabIndex        =   2
      Top             =   150
      Width           =   1380
   End
   Begin Sicmact.FlexEdit fgInteres 
      Height          =   3435
      Left            =   90
      TabIndex        =   13
      Top             =   630
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6059
      Cols0           =   24
      HighLight       =   1
      EncabezadosNombres=   $"frmAdeudPagLote.frx":030A
      EncabezadosAnchos=   "350-500-1600-1500-550-0-0-1200-0-1100-0-1100-0-0-0-0-1000-0-0-0-1100-700-0-1300"
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-X-11-X-X-X-X-16-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-C-L-L-C-C-C-R-R-C-R-R-R-R-C-C-R-C-C-L-C-R-L-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-2-2-2-2-2-2-0-0-2-0-0-0-0-2-0-0"
      TextArray0      =   "#"
      SelectionMode   =   1
      lbEditarFlex    =   -1  'True
      Appearance      =   0
      ColWidth0       =   345
      RowHeight0      =   300
   End
   Begin VB.Frame fraTransferencia 
      Caption         =   "Entidad Financiera"
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
      Height          =   1185
      Left            =   90
      TabIndex        =   6
      Top             =   4110
      Width           =   10095
      Begin Sicmact.EditMoney txtBancoImporte 
         Height          =   255
         Left            =   8130
         TabIndex        =   7
         Top             =   750
         Width           =   1785
         _ExtentX        =   2937
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BorderStyle     =   0
      End
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   360
         Left            =   1095
         TabIndex        =   8
         Top             =   300
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   635
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   6840
         Top             =   720
         Width           =   3105
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7050
         TabIndex        =   12
         Top             =   780
         Width           =   615
      End
      Begin VB.Label lblDesCtaIfTransf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1095
         TabIndex        =   11
         Top             =   720
         Width           =   5370
      End
      Begin VB.Label lblDescIfTransf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4155
         TabIndex        =   10
         Top             =   300
         Width           =   5790
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta N° :"
         Height          =   210
         Left            =   180
         TabIndex        =   9
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.TextBox txtTasaVac 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   5250
      TabIndex        =   1
      Top             =   150
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8730
      TabIndex        =   4
      Top             =   5385
      Width           =   1440
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   7305
      TabIndex        =   3
      Top             =   5385
      Width           =   1440
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   390
      Top             =   5340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   42
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdeudPagLote.frx":03DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdeudPagLote.frx":0BAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Sicmact.FlexEdit fgDetalle 
      Height          =   1365
      Left            =   90
      TabIndex        =   18
      Top             =   5340
      Visible         =   0   'False
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   2408
      Cols0           =   7
      HighLight       =   1
      EncabezadosNombres=   "#-Cuenta-Descripcion-Monto-Pos-Objeto-Monto VAC"
      EncabezadosAnchos=   "350-1200-2000-1200-0-0-1200"
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-3-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-R-L-L-R"
      FormatosEdit    =   "0-0-0-2-0-0-2"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   345
      RowHeight0      =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tasa VAC :"
      Height          =   195
      Left            =   4350
      TabIndex        =   5
      Top             =   225
      Width           =   810
   End
End
Attribute VB_Name = "frmAdeudPagLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lMN As Boolean
Dim lsCtaContDebe() As String
Dim lsCtaContHaber() As String
Dim aObj() As String
Dim lbCargar As Boolean
Dim lsGridDH As String
Dim lsPosCtaBusqueda As String
Dim lnTasaVac As Double
Dim lsCtaConcesional As String
Dim oAdeud As DCaja_Adeudados
Dim oOpe As DOperacion
Dim oCtaIf As NCajaCtaIF
Dim lsCtaOrdenD As String
Dim lsCtaOrdenH As String

'Efectivo
Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset

'Documento de Transferencia
Dim lsDocumento As String
Dim lnTpoDoc As TpoDoc
Dim lsNroDoc As String
Dim lsNroVoucher As String
Dim ldFechaDoc  As Date

'Variable para Refrescar Cuentas a Utilizar
Dim lsIFTpo As String
Dim lsPersCod As String

Private Function ValidaDatos() As Boolean
Dim lbMontoDet As Boolean
Dim i As Integer
ValidaDatos = False
    If fgInteres.TextMatrix(1, 0) = "" Then
        MsgBox "No existen Cuentas de para realizar la Operación", vbInformation, "Aviso"
        Me.cmdSalir.SetFocus
        Exit Function
    End If
    If nVal(txtBancoImporte) = 0 Then
        MsgBox "No se seleccionaro Pagarés a pagar", vbInformation, "¡Aviso!"
        fgInteres.SetFocus
        Exit Function
    End If
    If txtBuscaEntidad = "" Then
        MsgBox "Cuenta de Banco no Válida", vbInformation, "Aviso"
        txtBuscaEntidad.SetFocus
        Exit Function
    End If
    lbMontoDet = False
ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
Dim oDocPago As clsDocPago
Dim lsCuentaAho As String
Dim N           As Integer

Dim lsMovNro As String
Dim oCon     As NContFunciones
Dim oCaja As nCajaGeneral
Dim rsAdeud  As ADODB.Recordset

On Error GoTo AceptarErr
If Not ValidaDatos() Then
    Exit Sub
End If

Set oCon = New NContFunciones
Set oCaja = New nCajaGeneral

If MsgBox(" ¿ Desea Grabar Operación ? ", vbYesNo + vbQuestion, "Confirmación") = vbYes Then
    lsMovNro = oCon.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
    gsGlosa = gsOpeDesc
    oCaja.GrabaPagoCuotaAdeudadoslote lsMovNro, gsOpeCod, txtFecha, gsGlosa, fgInteres.GetRsNew(), _
 _
            gdFecSis, txtBuscaEntidad, _
 _
            fgDetalle.GetRsNew, _
            nVal(txtTasaVac), lsCtaConcesional, lsCtaOrdenD, lsCtaOrdenH
'    ImprimeAsientoContable lsMovNro, lsNroVoucher, lnTpoDoc, lsDocumento, True, False
    If MsgBox("Desea Realizar otra operación de Pago ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        For N = 1 To fgInteres.Rows - 1
            fgInteres.TextMatrix(N, 1) = ""
        Next
        txtBuscaEntidad = ""
        txtBancoImporte = "0.00"
        lblDescIfTransf = ""
        lblDesCtaIfTransf = ""
    Else
        Unload Me
    End If
End If
    
Exit Sub
AceptarErr:
    MsgBox "Error N° [" & Err.Number & "] " & TextErr(Err.Description), vbInformation, "Aviso"
End Sub

Private Sub cmdCalcular_Click()
lsIFTpo = ""
    If ValFecha(Me.txtFecha) = False Then Exit Sub
    CargaBancos CDate(txtFecha)
    fgInteres.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fgInteres_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
If Me.fgInteres.TextMatrix(1, 0) <> "" Then
    If fgInteres.TextMatrix(pnRow, pnCol) = "." Then
        If lsPersCod <> "" And lsPersCod <> Mid(fgInteres.TextMatrix(pnRow, 19), 4, 13) Then
            MsgBox "Debe seleccionar Pagares de una misma Institución para Pagar en Lote", vbInformation, "¡Aviso!"
            fgInteres.TextMatrix(pnRow, pnCol) = "0"
        Else
            If lsPersCod = "" Then
                lsIFTpo = Left(fgInteres.TextMatrix(fgInteres.Row, 19), 2)
                CargaCuentasGrid fgInteres.TextMatrix(fgInteres.Row, 19)
                lsPersCod = Mid(fgInteres.TextMatrix(pnRow, 19), 4, 13)
            End If
            fgInteres.TextMatrix(pnRow, 23) = Format(nVal(fgInteres.TextMatrix(fgInteres.Row, 7)) + nVal(fgInteres.TextMatrix(fgInteres.Row, 11)) + nVal(fgInteres.TextMatrix(fgInteres.Row, 9)) + nVal(fgInteres.TextMatrix(fgInteres.Row, 16)), gsFormatoNumeroView)
        End If
    Else
        fgInteres.TextMatrix(pnRow, 23) = ""
    End If
    txtBancoImporte = Format(fgInteres.SumaRow(23), gsFormatoNumeroView)
    If fgInteres.TextMatrix(pnRow, pnCol) = "" Then
        If nVal(txtBancoImporte) = 0 Then
            lsPersCod = ""
        End If
    End If
End If
End Sub

Private Sub Form_Activate()
    If lbCargar = False Then
        Unload Me
    End If
End Sub

Public Sub Inicio(psGridDH As String, psPosCtaBusqueda As String)
    lsGridDH = psGridDH
    lsPosCtaBusqueda = psPosCtaBusqueda
    Me.Show 1
End Sub

Private Sub Form_Load()
    Dim N As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer, J As Integer

    CentraForm Me
    Me.Caption = gsOpeDesc
    gsSimbolo = gcMN
    If Mid(gsOpeCod, 3, 1) = "2" Then
        gsSimbolo = gcME
    End If

    lbCargar = True
    Set oOpe = New DOperacion
    Set oAdeud = New DCaja_Adeudados
    Set oCtaIf = New NCajaCtaIF


    lnTasaVac = oAdeud.CargaIndiceVAC(gdFecSis)
    txtTasaVac = lnTasaVac

    txtOpeCod = gsOpeCod
    txtFecha.Text = Format(gdFecSis, gsFormatoFechaView)
    txtBuscaEntidad.rs = oOpe.GetOpeObj(gsOpeCod, "2")
    Set rs = oOpe.CargaOpeCta(gsOpeCod, "H", "6")
    If Not rs.EOF Then
        lsCtaConcesional = rs!cCtaContCod
    End If
    RSClose rs
    
    lsCtaOrdenD = oOpe.EmiteOpeCta(gsOpeCod, "D", 7)
    lsCtaOrdenH = oOpe.EmiteOpeCta(gsOpeCod, "H", 7)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oAdeud = Nothing
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
lblDescIfTransf = oCtaIf.NombreIF(Mid(txtBuscaEntidad.Text, 4, 14))
lblDesCtaIfTransf = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
cmdAceptar.SetFocus
End Sub

Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ValFecha(txtFecha) = False Then
            txtFecha.SetFocus
        Else
            cmdCalcular.SetFocus
            lnTasaVac = oAdeud.CargaIndiceVAC(txtFecha)
            txtTasaVac = lnTasaVac
        End If
    End If
End Sub

Private Sub CargaBancos(ldFecha As Date)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim N As Integer
    Dim lnMontoTotal As Currency
    Dim lnInteres As Currency
    Dim lnTotal As Integer, i As Integer
    Dim lnCapital As Currency

    lnTasaVac = oAdeud.CargaIndiceVAC(ldFecha)
    If lnTasaVac = 0 Then
        If MsgBox("Tasa VAC no ha sido definida para la fecha Ingresada" & Chr(10) & "Desea Proseguir con al Operación??", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    txtTasaVac = lnTasaVac
    
    'CARGAMOS LOS ADEUDADOS PENDIENTES
    Set rs = oAdeud.GetAdeudadosProvision(gsOpeCod, ldFecha, Mid(gsOpeCod, 3, 1))

    lnTotal = rs.RecordCount
    i = 0
    fgInteres.Rows = 2
    fgInteres.Clear
    fgInteres.FormaCabecera
    
    Do While Not rs.EOF
        i = i + 1
        fgInteres.AdicionaFila , , True
        N = fgInteres.Row
        fgInteres.TextMatrix(N, 2) = Trim(rs!cPersNombre)   'entidad
        fgInteres.TextMatrix(N, 3) = Trim(rs!cCtaIFDesc)    'cuenta
        fgInteres.TextMatrix(N, 4) = Trim(rs!nNroCuota)    ' numero de cuota pendiente
        'se oculta *
        fgInteres.TextMatrix(N, 5) = Format(rs!nSaldoCap, "#,#0.00") 'Saldocapital
        lnCapital = rs!nCapitalCuota
        'se oculta *
        fgInteres.TextMatrix(N, 6) = Format(lnCapital, "#,#0.00")  ' Saldo de Capital Base
        
        'Se muestra
        If rs!cMonedaPago = "2" And Mid(rs!cCtaIFCod, 3, 1) = "1" Then
            fgInteres.TextMatrix(N, 7) = Format(lnCapital * lnTasaVac, "#,#0.00") ' Saldo * la tasa vac
        Else
            fgInteres.TextMatrix(N, 7) = Format(lnCapital, "#,#0.00")  ' Saldo de Capital Normal
        End If
        'se oculta *

        fgInteres.TextMatrix(N, 8) = Format(rs!nInteresPagado, "#,#0.00") ' Interes acumulado pagado por cuota
        fgInteres.TextMatrix(N, 9) = Format(rs!nInteresPagado, "#,#0.00")
        
        If Val(rs!cIFTpo) = gTpoIFFuenteFinanciamiento Then
            lnMontoTotal = rs!nSaldoCap - rs!nSaldoConcesion
        Else
            lnMontoTotal = rs!nSaldoCap + rs!nInteresPagado
        End If
        lnInteres = oAdeud.CalculaInteres(rs!nDiasUltPAgo, rs!nPeriodo, rs!nInteres, lnMontoTotal)
        
        'se oculta
        fgInteres.TextMatrix(N, 10) = Format(lnInteres, "#,#0.00")
        'se muestra
        If rs!cMonedaPago = "2" And Mid(rs!cCtaIFCod, 3, 1) = "1" Then
            lnInteres = lnInteres * lnTasaVac
        End If
        fgInteres.TextMatrix(N, 11) = Format(lnInteres, "#,#0.00")
        'se oculta *

        fgInteres.TextMatrix(N, 12) = Format(rs!nInteresPagado + lnInteres, "#,#0.00")
        fgInteres.TextMatrix(N, 13) = Format((rs!nInteresPagado + lnInteres), "#0.00")
        fgInteres.TextMatrix(N, 14) = Format(rs!dCuotaUltPago, "dd/mm/yyyy")
        fgInteres.TextMatrix(N, 15) = Trim(rs!nPeriodo)
        fgInteres.TextMatrix(N, 16) = rs!nComision
        fgInteres.TextMatrix(N, 17) = Trim(rs!nDiasUltPAgo)
        fgInteres.TextMatrix(N, 18) = Trim(rs!cMonedaPago)
        fgInteres.TextMatrix(N, 19) = Trim(rs!cIFTpo & "." & rs!cPersCod & "." & rs!cCtaIFCod)
        fgInteres.TextMatrix(N, 20) = rs!dVencimiento
        fgInteres.TextMatrix(N, 21) = Trim(rs!nInteres)
        fgInteres.TextMatrix(N, 22) = rs!nSaldoCapLP
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Function Total(lbTotal As Boolean) As Currency
    Dim i As Integer
    Dim lnTotal As Currency
    If lbTotal = False Then
        For i = 1 To Me.fgDetalle.Rows - 1
            fgDetalle.TextMatrix(i, 5) = Trim(fgInteres.TextMatrix(fgInteres.Row, 19))
            Select Case Trim(fgDetalle.TextMatrix(i, 4))
                Case "0"
                   fgDetalle.TextMatrix(i, 3) = Format(fgInteres.TextMatrix(fgInteres.Row, 7), "#0.00")  'capital que se muestra
                    fgDetalle.TextMatrix(i, 6) = Format(fgInteres.TextMatrix(fgInteres.Row, 6), "#0.00")  'capital que se oculta
                Case "1"
                    fgDetalle.TextMatrix(i, 3) = Format(fgInteres.TextMatrix(fgInteres.Row, 11), "#0.00")
                    fgDetalle.TextMatrix(i, 6) = Format(fgInteres.TextMatrix(fgInteres.Row, 10), "#0.00")
                Case "2"
                    fgDetalle.TextMatrix(i, 3) = Format(fgInteres.TextMatrix(fgInteres.Row, 9), "#0.00")
                    fgDetalle.TextMatrix(i, 6) = Format(fgInteres.TextMatrix(fgInteres.Row, 8), "#0.00")
                Case "3"  'Comision
                    fgDetalle.TextMatrix(i, 3) = Format(fgInteres.TextMatrix(fgInteres.Row, 16), "#0.00")
                    fgDetalle.TextMatrix(i, 6) = Format(fgInteres.TextMatrix(fgInteres.Row, 16), "#0.00")
            End Select
        Next
    End If
    lnTotal = 0
    For i = 1 To fgDetalle.Rows - 1
        lnTotal = lnTotal + CCur(IIf(fgDetalle.TextMatrix(i, 3) = "", "0", fgDetalle.TextMatrix(i, 3)))
    Next
'    lblTotal = Format(lnTotal, "#,#0.00")
End Function

Private Sub CargaCuentasGrid(psIFCod As String)
Dim rs As ADODB.Recordset
Dim oOpe As New DOperacion
If Not psIFCod = "" Then
    Set rs = oOpe.CargaOpeCtaIF(gsOpeCod, psIFCod, "D")
    If Not RSVacio(rs) Then
        fgDetalle.Rows = 2
        fgDetalle.Clear
        fgDetalle.FormaCabecera
        Do While Not rs.EOF
            fgDetalle.AdicionaFila
            fgDetalle.TextMatrix(fgDetalle.Row, 1) = rs!cCtaContCod
            fgDetalle.TextMatrix(fgDetalle.Row, 2) = Trim(rs!cCtaContDesc)
            fgDetalle.TextMatrix(fgDetalle.Row, 4) = Trim(rs!cOpeCtaOrden)
            rs.MoveNext
        Loop
    Else
        lbCargar = False
        MsgBox "No se han definido Cuentas Contables para Operación", vbInformation, "Aviso"
    End If
End If
    fgDetalle.TopRow = 1
    fgDetalle.Row = 1
    RSClose rs
    Set oOpe = Nothing
End Sub

Private Sub CalculoInteres(ByVal lnDias As Long)
    Dim lnPeriodo As Long
    Dim lnTasaInt As Currency
    Dim lnMontoTotal As Currency
    Dim lnInteres As Currency

    fgInteres.TextMatrix(fgInteres.Row, 17) = lnDias
    If Val(Mid(fgInteres.TextMatrix(fgInteres.Row, 19), 1, 2)) = gTpoIFFuenteFinanciamiento Then
        lnMontoTotal = CCur(fgInteres.TextMatrix(fgInteres.Row, 4))
    Else
        lnMontoTotal = CCur(fgInteres.TextMatrix(fgInteres.Row, 4)) + CCur(fgInteres.TextMatrix(fgInteres.Row, 8))
    End If
    lnTasaInt = CCur(fgInteres.TextMatrix(fgInteres.Row, 21))
    lnPeriodo = Val(fgInteres.TextMatrix(fgInteres.Row, 15))
    lnInteres = oAdeud.CalculaInteres(lnDias, lnPeriodo, lnTasaInt, lnMontoTotal)

    fgInteres.TextMatrix(fgInteres.Row, 10) = Format(lnInteres, "#0.00")
    If fgInteres.TextMatrix(fgInteres.Row, 18) = "2" And Mid(fgInteres.TextMatrix(fgInteres.Row, 19), 9, 1) = "1" Then
        fgInteres.TextMatrix(fgInteres.Row, 11) = Format(lnInteres * lnTasaVac, "#0.00")
    Else
        fgInteres.TextMatrix(fgInteres.Row, 11) = Format(lnInteres, "#0.00")
    End If
End Sub

