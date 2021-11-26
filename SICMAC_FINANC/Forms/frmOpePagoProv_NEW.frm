VERSION 5.00
Begin VB.Form frmOpePagoProv_NEW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago a Proveedores"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12810
   Icon            =   "frmOpePagoProv_NEW.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   320
      Left            =   11640
      TabIndex        =   28
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   320
      Left            =   10537
      TabIndex        =   8
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdEmitir 
      Caption         =   "&Emitir"
      Height          =   320
      Left            =   9440
      TabIndex        =   7
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame fraEntidadPagadora 
      Caption         =   "Entidad Pagadora"
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
      Height          =   625
      Left            =   80
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   9015
      Begin Sicmact.TxtBuscar txtEntidadPagadoraCod 
         Height          =   300
         Left            =   150
         TabIndex        =   6
         Top             =   210
         Width           =   2570
         _ExtentX        =   4524
         _ExtentY        =   529
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblEntidadPagadoraNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2760
         TabIndex        =   23
         Top             =   210
         Width           =   6075
      End
   End
   Begin VB.Frame fraEntidad 
      Caption         =   "Observaciones"
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
      Height          =   660
      Left            =   80
      TabIndex        =   21
      Top             =   5160
      Width           =   12615
      Begin VB.TextBox txtObservaciones 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   12375
      End
   End
   Begin VB.CheckBox chkAfectoITF 
      Caption         =   "Afecto a ITF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11040
      TabIndex        =   4
      Top             =   4920
      Value           =   1  'Checked
      Width           =   1605
   End
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      Caption         =   "Todos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   915
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
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
      Height          =   700
      Left            =   80
      TabIndex        =   9
      Top             =   0
      Width           =   12615
      Begin VB.TextBox txtAnio 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   29
         Top             =   240
         Width           =   480
      End
      Begin VB.Frame fraEntidadReceptora 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5880
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   5415
         Begin Sicmact.TxtBuscar txtEntidadReceptoraCod 
            Height          =   300
            Left            =   600
            TabIndex        =   25
            Top             =   150
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            sTitulo         =   ""
         End
         Begin VB.Label lblEntidadReceptoraNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2040
            TabIndex        =   27
            Top             =   150
            Width           =   3315
         End
         Begin VB.Label lblEntidadReceptora 
            Caption         =   "Entidad:"
            Height          =   255
            Left            =   0
            TabIndex        =   26
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   320
         Left            =   11400
         TabIndex        =   2
         Top             =   270
         Width           =   1095
      End
      Begin VB.ComboBox cboTpoPago 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2100
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1380
      End
      Begin VB.Frame FraTipoB 
         Height          =   885
         Left            =   2700
         TabIndex        =   10
         Top             =   1005
         Width           =   5295
         Begin VB.OptionButton optAge 
            Appearance      =   0  'Flat
            Caption         =   "Logistica"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   13
            Top             =   165
            Width           =   1125
         End
         Begin VB.OptionButton optAge 
            Appearance      =   0  'Flat
            Caption         =   "x Agencia"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   12
            Top             =   165
            Width           =   1065
         End
         Begin VB.OptionButton optAge 
            Appearance      =   0  'Flat
            Caption         =   "Todos"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   165
            Value           =   -1  'True
            Width           =   855
         End
         Begin Sicmact.TxtBuscar txtAge 
            Height          =   345
            Left            =   930
            TabIndex        =   14
            Top             =   450
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   609
            Appearance      =   0
            BackColor       =   14811132
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            sTitulo         =   ""
            EnabledText     =   0   'False
         End
         Begin VB.Label lblAgencia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1950
            TabIndex        =   16
            Top             =   450
            Width           =   3255
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Agencia:"
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
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   510
            Width           =   765
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Pago:"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1875
         TabIndex        =   18
         Top             =   270
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Mes:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   285
         Width           =   375
      End
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   3885
      Left            =   75
      TabIndex        =   3
      Top             =   990
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   6853
      Cols0           =   32
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmOpePagoProv_NEW.frx":030A
      EncabezadosAnchos=   "400-0-500-2100-1140-4000-0-1250-0-0-0-0-0-0-1400-1500-2000-2300-1500-1800-1200-2500-2500-0-0-0-0-0-0-0-0-0"
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
      ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-C-L-L-R-L-C-C-C-C-C-L-R-R-R-R-L-R-R-R-R-C-C-R-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-0-0-0-0-0-0-0-2-0-2-2-0-2-2-2-4-4-4-3-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmOpePagoProv_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************************************
'** Nombre : frmOpePagoProv_NEW
'** Descripción : Formulario para el nuevo pago de Comprobantes segun ERS062-2013 basado en el formulario frmOpePagProv
'** Creación : EJVG, 20131118 11:00:00 AM
'**********************************************************************************************************************
Option Explicit
Dim lsCtaContDebeB As String
Dim lsCtaContDebeS As String
Dim lsCtaContDebeRH As String
Dim lsCtaContDebeRHJ As String
Dim lsCtaContDebeSegu As String
Dim lsCtaContDebePagoVarios As String
Dim lsCtaContDebeBLeasingMN As String
Dim lsCtaContDebeBLeasingME As String
Dim lsCtaContDebeBLeasingSMN As String
Dim lsCtaContDebeBLeasingSME As String
Dim lsCtaContHaberFepMacS As String 'VAPA20170724
Dim lsCtaITFD As String
Dim lsCtaITFH As String
Dim lsDocs As String
Dim lsTipoB As String
Dim lnTipoPago As Integer
Dim fMatProveedor() As String
Dim fsPersCodCMACMaynas As String
Dim fsPersCodBCP As String
Dim lsDocTpo As TpoDoc
Dim lmn As Boolean
Dim lsFileCarta As String
Dim lsCtaCore As String
Dim lsCtaSAF As String
Dim fsCtaContPenalidadH As String
Dim lsCtaContProvNewApert As String 'NAGL INC1712260008 20171227
Dim lsCtaContProvNewApertFull As String 'NAGL INC1712260008 20171227

Dim fnTpoPago As LogTipoPagoComprobante
Private Type PagoProv
    lnItem As Integer
    lsMovNro As String
    lsVoucherPago As String
End Type
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAnio.SetFocus
    End If
End Sub
Private Sub cboTpoPago_Click()
    Dim oOpe As New DOperacion
    fnTpoPago = 0
    
    fraEntidadReceptora.Visible = False
    fraEntidadPagadora.Visible = False
    If cboTpoPago.ListIndex <> -1 Then
        fnTpoPago = CInt(Trim(Right(cboTpoPago.Text, 3)))
    End If
    If fnTpoPago = gPagoTransferencia Then
        fraEntidadPagadora.Visible = True
        txtEntidadPagadoraCod.rs = oOpe.GetRsOpeObj(gsOpeCod, "1", , , , fsPersCodBCP)
    ElseIf fnTpoPago = gPagoCheque Then
        fraEntidadReceptora.Visible = True
        fraEntidadPagadora.Visible = True
        txtEntidadPagadoraCod.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")
    End If
    Set oOpe = Nothing
End Sub
Private Sub cboTpoPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fnTpoPago = gPagoCheque Then
            If txtEntidadReceptoraCod.Visible And txtEntidadReceptoraCod.Enabled Then
                txtEntidadReceptoraCod.SetFocus
            End If
        Else
            cmdProcesar.SetFocus
        End If
    End If
End Sub

Private Sub chkTodos_Click()
    Dim frmSel As frmOpePagoProvSel
    Dim i As Integer
    Dim lsOpcion As String
    
    On Error GoTo ErrChkTodos
    
    fg.col = 1
    If fnTpoPago = 0 Then
        chkTodos.value = 0
        Exit Sub
    End If
    If chkTodos.value = 1 Then
        If fg.TextMatrix(1, 0) <> "" Then
            For i = 1 To fg.Rows - 1
                fg.TextMatrix(i, 2) = ""
            Next
            Set frmSel = New frmOpePagoProvSel
            lsOpcion = frmSel.inicio(fMatProveedor, fnTpoPago)
            Set frmSel = Nothing
            If lsOpcion <> "" Then
                If lsOpcion = "TODOS" Then
                    For i = 1 To fg.Rows - 1
                        fg.TextMatrix(i, 2) = "1"
                    Next
                Else
                    For i = 1 To fg.Rows - 1
                        If fg.TextMatrix(i, 8) = lsOpcion Then
                            fg.TextMatrix(i, 2) = "1"
                        End If
                    Next
                End If
            Else
                chkTodos.value = 0
            End If
        Else
            chkTodos.value = 0
        End If
    Else
        If fg.TextMatrix(1, 0) <> "" Then
            For i = 1 To fg.Rows - 1
                fg.TextMatrix(i, 2) = ""
            Next
        End If
    End If
    Exit Sub
ErrChkTodos:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdProcesar_Click()
    Dim oOpeL As DOperacion
    Dim oDCaja As DCajaGeneral
    Dim oTC As nTipoCambio
    Dim rs As ADODB.Recordset
    Dim bTieneDetra As Boolean
    Dim cCtaDetraTemp As String
    Dim nCantTempo As Integer
    Dim lsTipoInterfaz As String
    Dim lnTCFijo As Currency
    Dim lsCtaContDebeBME As String
    Dim lsCtaContDebeSME As String
    Dim lsCtaContDebeRHME As String
    Dim lsCtaContDebeRHJME As String
    Dim lsCtaContDebeSeguME As String
    Dim lsCtaContDebePagoVariosME As String
    Dim lsProveedorLeasingMN As String
    Dim lsProveedorLeasingME As String
    Dim lsProveedorLeasingSMN As String
    Dim lsProveedorLeasingSME As String
    Dim lsListaProveedoresSinIFi As String
    Dim lnMes As Integer, lnTpoPago As Integer, Index As Integer, iMat As Integer
    Dim nItem As Long
    Dim ldFecha As Date
    Dim lsIFiCod As String
    Dim MatProveedor() As String
    Dim bExiste As Boolean
    
    On Error GoTo ErrCargaProveedores
    
    If Not validaProcesar Then Exit Sub
    
    Set oOpeL = New DOperacion
    Set oDCaja = New DCajaGeneral
    Set oTC = New nTipoCambio
    Set rs = New ADODB.Recordset
    Screen.MousePointer = 11
    
    chkTodos.value = 0
    FormateaFlex fg
    lnTCFijo = oTC.EmiteTipoCambio(gdFecSis, TCFijoMes)
    cCtaDetraTemp = Mid(cCtaDetraccionProvision, 1, 2) & Mid(gsOpeCod, 3, 1) & Mid(cCtaDetraccionProvision, 4, Len(cCtaDetraccionProvision) - 2)
    
    lsProveedorLeasingMN = oOpeL.EmiteOpeCtaLeasing("421110", 1)
    lsProveedorLeasingME = oOpeL.EmiteOpeCtaLeasing("422110", 1)
    lsProveedorLeasingSMN = oOpeL.EmiteOpeCtaLeasing("421110", 2)
    lsProveedorLeasingSME = oOpeL.EmiteOpeCtaLeasing("422110", 2)
    
    Set oTC = Nothing
    Set oOpeL = Nothing
    
    lsTipoInterfaz = "PAGO"
    lnMes = CInt(Trim(Right(cboMes.Text, 2)))
    ldFecha = CDate("01/" & Format(lnMes, "00") & "/" & Val(txtAnio.Text))
    lnTpoPago = CInt(Trim(Right(cboTpoPago.Text, 2)))
    If lnTpoPago = 1 Then
        lsIFiCod = fsPersCodCMACMaynas
    ElseIf lnTpoPago = 2 Then
        lsIFiCod = fsPersCodBCP
    ElseIf lnTpoPago = 3 Then
        lsIFiCod = txtEntidadReceptoraCod.Text
    End If
    
    'Set rs = oDCaja.GetDatosProvisionesProveedores("'" & lsProveedorLeasingSMN & "','" & lsProveedorLeasingSME & "','" & lsProveedorLeasingMN & "','" & lsProveedorLeasingME & "','" & lsCtaContDebeB & "','" & lsCtaContDebeS & "','" & lsCtaContDebeRH & "','" & lsCtaContDebeRHJ & "', '" & lsCtaContDebeSegu & "', '" & lsCtaContDebePagoVarios & "' ", lsDocs, CDate("01/01/1900"), CDate("01/01/1900"), , 3, cCtaDetraTemp, "'" & lsCtaContDebeBME & "','" & lsCtaContDebeSME & "','" & lsCtaContDebeRHME & "','" & lsCtaContDebeRHJME & "', '" & lsCtaContDebeSeguME & "', '" & lsCtaContDebePagoVariosME & "'", lsTipoInterfaz, "", "", False, "", lsTipoB, lnTipoPago, True, ldFecha, lsIFiCod)
    lsCtaContProvNewApertFull = oDCaja.GetProveedoresCtasAperturadasSBS(gsOpeCod) 'NAGL INC1712260008
    
    If lsCtaContProvNewApertFull <> "" Then
        Set rs = oDCaja.GetDatosProvisionesProveedores("'" & IIf(Mid(gsOpeCod, 3, 1) = "1", lsProveedorLeasingSMN, lsProveedorLeasingSME) & "','" & IIf(Mid(gsOpeCod, 3, 1) = "1", lsProveedorLeasingMN, lsProveedorLeasingME) & "','" & lsCtaContDebeB & "','" & lsCtaContDebeS & "','" & lsCtaContDebeRH & "','" & lsCtaContDebeRHJ & "', '" & lsCtaContDebeSegu & "', '" & lsCtaContDebePagoVarios & "','" & lsCtaContHaberFepMacS & "'," & lsCtaContProvNewApertFull & "", lsDocs, CDate("01/01/1900"), CDate("01/01/1900"), , 3, cCtaDetraTemp, "'" & lsCtaContDebeBME & "','" & lsCtaContDebeSME & "','" & lsCtaContDebeRHME & "','" & lsCtaContDebeRHJME & "', '" & lsCtaContDebeSeguME & "', '" & lsCtaContDebePagoVariosME & "'", lsTipoInterfaz, "", "", False, "", lsTipoB, lnTipoPago, True, ldFecha, lsIFiCod)
    Else
        Set rs = oDCaja.GetDatosProvisionesProveedores("'" & IIf(Mid(gsOpeCod, 3, 1) = "1", lsProveedorLeasingSMN, lsProveedorLeasingSME) & "','" & IIf(Mid(gsOpeCod, 3, 1) = "1", lsProveedorLeasingMN, lsProveedorLeasingME) & "','" & lsCtaContDebeB & "','" & lsCtaContDebeS & "','" & lsCtaContDebeRH & "','" & lsCtaContDebeRHJ & "', '" & lsCtaContDebeSegu & "', '" & lsCtaContDebePagoVarios & "','" & lsCtaContHaberFepMacS & "' ", lsDocs, CDate("01/01/1900"), CDate("01/01/1900"), , 3, cCtaDetraTemp, "'" & lsCtaContDebeBME & "','" & lsCtaContDebeSME & "','" & lsCtaContDebeRHME & "','" & lsCtaContDebeRHJME & "', '" & lsCtaContDebeSeguME & "', '" & lsCtaContDebePagoVariosME & "'", lsTipoInterfaz, "", "", False, "", lsTipoB, lnTipoPago, True, ldFecha, lsIFiCod) 'EJVG20140416 VAPA20170724 AGREGO lsCtaContHaberFepMacS
    End If 'NAGL INC1712260008 'NAGL Agregó lsProveedorCtasApertSBS, en este método según INC1712260008 y Condicional
    
    
    Set oDCaja = Nothing

    If rs.EOF Then
        Screen.MousePointer = 0
        Set rs = Nothing
        MsgBox "No existen Comprobantes Pendientes", vbInformation, "Aviso"
        Exit Sub
    End If
    
    nCantTempo = 0
    ReDim MatProveedor(1, 0)
    ReDim fMatProveedor(1, 0)
    Do While Not rs.EOF
        If rs!cPersCodIFI <> "" Then
            fg.AdicionaFila
            nItem = fg.Row
    
            fg.TextMatrix(nItem, 1) = nItem
            fg.TextMatrix(nItem, 3) = Mid(rs!cDocAbrev & space(3), 1, 3) & " " & rs!cDocNro
            fg.TextMatrix(nItem, 4) = rs!dDocFecha
            fg.TextMatrix(nItem, 5) = PstaNombre(rs!cPersona, True)
            fg.TextMatrix(nItem, 6) = rs!cMovDesc
            fg.TextMatrix(nItem, 7) = Format(rs!nMovImporte, gsFormatoNumeroView)
            fg.TextMatrix(nItem, 8) = rs!cPersCod
            fg.TextMatrix(nItem, 9) = rs!cMovNro
            fg.TextMatrix(nItem, 10) = rs!nMovNro
            fg.TextMatrix(nItem, 11) = rs!nDocTpo
            fg.TextMatrix(nItem, 12) = rs!cDocNro
            fg.TextMatrix(nItem, 13) = rs!cCtaContCod
            fg.TextMatrix(nItem, 14) = GetFechaMov(rs!cMovNro, True)
            fg.TextMatrix(nItem, 24) = rs!Movenvio
            fg.TextMatrix(nItem, 25) = rs!Agencia
            fg.TextMatrix(nItem, 26) = rs!nPenalidad
            fg.TextMatrix(nItem, 17) = Format(rs!nimportecoactivo, gsFormatoNumeroView)
            fg.TextMatrix(nItem, 18) = Format(rs!montopago, gsFormatoNumeroView)
            If rs!nMovImporteSoles <> rs!nMovImporte Then
                fg.TextMatrix(nItem, 15) = Format(Round((rs!nMovImporte - rs!MontoPagadoSUNAT) * rs!nTpoCambio, 2), gsFormatoNumeroView)
                fg.TextMatrix(nItem, 20) = Format(Round((rs!nMovImporte - rs!MontoPagadoSUNAT) * lnTCFijo, 2) - Round(CDbl(fg.TextMatrix(nItem, 18)) * lnTCFijo, 2) - Round(Round(rs!nimportecoactivo / rs!nTpoCambio, 2) * lnTCFijo, 2), "0.00")
                fg.TextMatrix(nItem, 21) = Format(rs!MontoPagadoSUNAT, gsFormatoNumeroView)
                fg.TextMatrix(nItem, 22) = Format(rs!MontoPagadoSUNATS, gsFormatoNumeroView)
            Else
                fg.TextMatrix(nItem, 15) = Format(rs!nMovImporteSoles - rs!MontoPagadoSUNAT, gsFormatoNumeroView)
                fg.TextMatrix(nItem, 20) = "0.00"
            End If
            fg.TextMatrix(nItem, 19) = IIf(rs!nMovImporteSoles = rs!nMovImporte, "SOLES", "DOLARES")
            fg.TextMatrix(nItem, 27) = rs!cPersCodIFI
            fg.TextMatrix(nItem, 28) = rs!cPersNombreIFi
            fg.TextMatrix(nItem, 29) = rs!cCtaCodIFi
            
            bExiste = False
            If UBound(fMatProveedor, 2) >= 1 Then
                For iMat = 1 To UBound(fMatProveedor, 2)
                    If fMatProveedor(0, iMat) = fg.TextMatrix(nItem, 8) Then
                        bExiste = True
                        Exit For
                    End If
                Next
            End If
            If Not bExiste Then
                Index = UBound(fMatProveedor, 2) + 1
                ReDim Preserve fMatProveedor(1, Index)
                fMatProveedor(0, Index) = fg.TextMatrix(nItem, 8)
                fMatProveedor(1, Index) = fg.TextMatrix(nItem, 5)
            End If
        Else
            bExiste = False
            If UBound(MatProveedor, 2) >= 1 Then
                For iMat = 1 To UBound(MatProveedor, 2)
                    If MatProveedor(0, iMat) = rs!cPersCod Then
                        bExiste = True
                        Exit For
                    End If
                Next
            End If
            If Not bExiste Then
                Index = UBound(MatProveedor, 2) + 1
                ReDim Preserve MatProveedor(1, Index)
                MatProveedor(0, Index) = rs!cPersCod
                MatProveedor(1, Index) = PstaNombre(rs!cPersona, True)
                lsListaProveedoresSinIFi = lsListaProveedoresSinIFi & "- " & MatProveedor(1, Index) & Chr(13)
            End If
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
    fg.TopRow = 1
    fg.Row = 1
    txtObservaciones.Text = fg.TextMatrix(fg.Row, 6)
    If fg.TextMatrix(1, 0) <> "" Then
        fraBusqueda.Enabled = False
    End If
    
    Screen.MousePointer = 0
    If Len(Trim(lsListaProveedoresSinIFi)) > 0 Then
        MsgBox "Los sgtes Proveedores no tienen configurado en el Módulo de Logistica sus cuentas en Instituciones Financieras:" & Chr(13) & Chr(13) & lsListaProveedoresSinIFi & Chr(13) & "Coordine con el Dpto. de Logistica el registro de los mismos.", vbInformation, "Aviso"
    End If
    Exit Sub
ErrCargaProveedores:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub
Private Sub cmdEmitir_Click()
    Dim oOpe As DOperacion
    Dim oDCtaIF As DCajaCtasIF
    Dim oCtasIF As NCajaCtaIF
    Dim oNContFunc As NContFunciones
    Dim oDocPago As clsDocPago
    Dim oNCaja As nCajaGeneral
    Dim oConst As NConstSistemas
    Dim oImpuesto As DImpuesto
    Dim oPrevio As clsPrevioFinan
    Dim rsBilletaje As ADODB.Recordset
    Dim oDCapta As DCapMantenimiento
    Dim oDoc As DDocumento
    Dim oImp As NContImprimir
    Dim oDis As NRHProcesosCierre
    
    Dim K As Integer
    Dim lsEntidadOrig As String
    Dim lsCtaEntidadOrig As String
    Dim lsPersNombre As String
    Dim lsCuentaAho As String
    Dim lnImporteB As Currency
    Dim lnImporteS As Currency
    Dim lnImporteAjusteDolaresS As Currency
    Dim lnImporteAjusteDolaresB As Currency
    Dim lnImporteAjusteDolaresRH As Currency
    Dim lnImporteAjusteDolaresRHJ As Currency
    Dim lnImporteAjusteDolaresSEGU As Currency
    Dim lnImporteAjusteDolaresPagoVarios As Currency
    Dim lsSubCuentaIF As String
    Dim lsPersCod As String
    Dim lsDocNRo As String
    Dim lsMovNro As String
    Dim lsOpeCod As String
    Dim lsCtaBanco As String
    Dim lsCtaContHaber As String
    Dim lsPersCodIF As String
    Dim lbEfectivo As Boolean
    Dim lsTpoIf As String
    Dim lsDocVoucher As String
    Dim lsFecha As String
    Dim lsGlosa As String 'PASIERS1242014
    Dim lnITFValor As Double
    
    Dim lnMontoDif As Currency
    Dim lsCtaDiferencia As String
    Dim nDocs As Integer
    Dim lsCabeImpre As String
    Dim lsImpre As String
    Dim lsCadBol As String
        
    Dim lbBitReten As Boolean
    Dim lbDocConIGV As Boolean
    Dim lsCtaReten As String
    Dim lbBCAR As Boolean
    Dim lnTasaImp As Currency
    Dim lnIngresos As Currency
    Dim lnRetencion As Currency
    Dim lnTopeRetencion As Currency
    Dim lnRetAct As Currency
    Dim lnRetActME As Currency
    Dim lsComprobante As String
    Dim lnDocProv As String
    Dim lsCtaContDebeRH As String
    Dim lsCtaContDebeRHJ As String
    Dim lsCtaContDebePagoVarios As String
    Dim lnImporteRH As Currency
    Dim lnImporteRHJ As Currency
    Dim lnImporteSEGU As Currency
    Dim lnImportePAGOVARIOS As Currency
    Dim lOk As Integer
    Dim lsCtaContLeasing As String
    Dim lnITF As Double
    Dim lnMontoPago As Currency
    Dim lsIFiCod As String, lsIFiNombre As String, lsIFiCtaCod As String
    Dim bSelecciono As Boolean
    Dim lsPlanillaNro As String
    Dim bAceptaRetencion As Boolean
    Dim MatPagos() As PagoProv
    Dim lsDocTpoTmp As String
    Dim lnNroPagos As Integer
    Dim lsDocNroTmp As String, lsDocVoucherTmp As String
    Dim lnMovNroProv As Long
    Dim rs As New ADODB.Recordset 'NAGL INC1712260008
    Dim oDCaja As New DCajaGeneral 'NAGL INC1712260008
    lsCtaContProvNewApert = "" 'NAGL INC1712260008
    Set oOpe = New DOperacion 'NAGL INC1712260008
    
    lsOpeCod = gsOpeCod
    lnTipoPago = 0
    lsDocTpo = "-1"
    lsCtaContDebeB = oOpe.EmiteOpeCta(gsOpeCod, "D", "0") 'NAGL 20171229
    
    cmdEmitir.Enabled = False 'PASI20150216
    
    If fnTpoPago = gPagoCuentaCMAC Then
        lsDocTpo = TpoDocNotaAbono
    ElseIf fnTpoPago = gPagoTransferencia Then
        lsDocTpo = TpoDocCarta
    ElseIf fnTpoPago = gPagoCheque Then
        lsDocTpo = TpoDocCheque
    End If

    On Error GoTo NoGrabo
    If ValidaInterfaz = False Then
        cmdEmitir.Enabled = True
        Exit Sub
    End If
    
    If lsDocTpo = "-1" Then
        MsgBox "El Sistema no ha podido predeterminar la forma de Pago", vbInformation, "Aviso"
        cmdEmitir.Enabled = True 'PASI20150216
        Exit Sub
    End If
    If lsDocTpo = TpoDocOrdenPago And chkAfectoITF.value = 0 Then ' Orden de pago
        MsgBox "Orden de Pago debe ser Afecto a ITF", vbInformation, "Aviso"
        cmdEmitir.Enabled = True 'PASI20150216
        Exit Sub
    End If
    If lsDocTpo = TpoDocNotaAbono And chkAfectoITF.value = 0 Then 'Abono en cuenta
        MsgBox "Abono en Cuenta debe ser Afecto a ITF", vbInformation, "Aviso"
        cmdEmitir.Enabled = True 'PASI20150216
        Exit Sub
    End If
    
    Set oDCapta = New DCapMantenimiento
    Set oOpe = New DOperacion
    For K = 1 To fg.Rows - 1
        If fg.TextMatrix(K, 2) = "." Then
            bSelecciono = True
            If fg.TextMatrix(K, 26) <> lnTipoPago Then
                MsgBox "No se puede realizar diferentes tipos de pagos", vbInformation, "¡Aviso!"
                fg.SetFocus
                cmdEmitir.Enabled = True 'PASI20150216
                Exit Sub
            End If
            If lsDocTpo = TpoDocCheque Then
                If lsPersCod <> "" Then
                    If lsPersCod <> fg.TextMatrix(K, 8) Then
                        MsgBox "No se puede hacer Pago a Proveedores diferentes cuando es con Cheque", vbInformation, "¡Aviso!"
                        fg.SetFocus
                        cmdEmitir.Enabled = True 'PASI20150216
                        Exit Sub
                    End If
                End If
            ElseIf lsDocTpo = TpoDocNotaAbono Then
                If Len(Trim(fg.TextMatrix(K, 27))) <> 13 Or Len(Trim(fg.TextMatrix(K, 29))) <> 18 Then
                    MsgBox "El proveedor " & fg.TextMatrix(K, 5) & " no tiene configurado la Cuenta de Ahorro en el Sistema", vbInformation, "Aviso"
                    fg.Row = K
                    fg.col = 1
                    fg.TopRow = K
                    fg.SetFocus
                    cmdEmitir.Enabled = True 'PASI20150216
                    Exit Sub
                End If
                'Verificamos la existencia de la cuenta del Proveedor
                If Not oDCapta.CuentaEsValida(fg.TextMatrix(K, 29), fg.TextMatrix(K, 8)) Then
                    MsgBox "La cuenta del proveedor " & fg.TextMatrix(K, 5) & " no es correcta o no esta vigente, coordine con el Dpto. de Logistica", vbInformation, "Aviso"
                    fg.Row = K
                    fg.col = 1
                    fg.TopRow = K
                    fg.SetFocus
                    Set oDCapta = Nothing
                    cmdEmitir.Enabled = True 'PASI20150216
                    Exit Sub
                End If
            End If
            
            'Se copió del bloque de abajo para sacar el monto total para el ingreso de cheque ***
            If lsCtaContDebeB = fg.TextMatrix(K, 13) And Not (lsCtaContDebeBLeasingMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingME = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSME = fg.TextMatrix(K, 13)) Then
                lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
            End If
            
            
            If lsCtaContDebeS = fg.TextMatrix(K, 13) Then
                lnImporteS = lnImporteS + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresS = lnImporteAjusteDolaresS + CCur(fg.TextMatrix(K, 20))
            End If
            If lsCtaContDebeRH = fg.TextMatrix(K, 13) Then
                lnImporteRH = lnImporteRH + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresRH = lnImporteAjusteDolaresRH + CCur(fg.TextMatrix(K, 20))
            End If
            If lsCtaContDebeRHJ = fg.TextMatrix(K, 13) Then
                lnImporteRHJ = lnImporteRHJ + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresRHJ = lnImporteAjusteDolaresRHJ + CCur(fg.TextMatrix(K, 20))
            End If
            If lsCtaContDebeSegu = fg.TextMatrix(K, 13) Then
                lnImporteSEGU = lnImporteSEGU + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresSEGU = lnImporteAjusteDolaresSEGU + CCur(fg.TextMatrix(K, 20))
            End If
            If lsCtaContDebePagoVarios = fg.TextMatrix(K, 13) Then
                lnImportePAGOVARIOS = lnImportePAGOVARIOS + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresPagoVarios = lnImporteAjusteDolaresPagoVarios + CCur(fg.TextMatrix(K, 20))
            End If
            If lsCtaContDebeBLeasingMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSMN = fg.TextMatrix(K, 13) Then
                lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
                lsCtaContLeasing = fg.TextMatrix(K, 13)
            End If
            If lsCtaContDebeBLeasingME = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSME = fg.TextMatrix(K, 13) Then
                lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
                lsCtaContLeasing = fg.TextMatrix(K, 13)
            End If
            '************************************************************************************
            lsOpeCod = oOpe.EmiteOpeDoc(Mid(gsOpeCod, 1, 5), lsDocTpo)
            If lsOpeCod = "" Then
                MsgBox "No se asignó Documentos de Referencia a Operación de Pago", vbInformation, "Aviso"
                cmdEmitir.Enabled = True 'PASI20150216
                Exit Sub
            End If
            lsCtaContHaber = oOpe.EmiteOpeCta(lsOpeCod, "H", , txtEntidadPagadoraCod.Text, ObjEntidadesFinancieras)
            If lsCtaContDebeB = "" Or lsCtaContDebeS = "" Or lsCtaContHaber = "" Then
                MsgBox "Cuentas Contables no determinadas Correctamente" & oImpresora.gPrnSaltoLinea & "consulte con Sistemas", vbInformation, "Aviso"
                cmdEmitir.Enabled = True 'PASI20150216
                Exit Sub
            End If
            If fsCtaContPenalidadH = "" Then
                MsgBox "Cuentas Contable Haber de Penalidad no determinada" & oImpresora.gPrnSaltoLinea & "consulte con Sistemas", vbInformation, "Aviso"
                cmdEmitir.Enabled = True 'PASI20150216
                Exit Sub
            End If
            lsCtaDiferencia = oOpe.EmiteOpeCta(lsOpeCod, "D", "2")
            
            lsPersCod = fg.TextMatrix(K, 8)
            lsPersNombre = fg.TextMatrix(K, 5)
        End If
    Next
    Set oDCapta = Nothing
    
    If Not bSelecciono Then
        MsgBox "No se seleccionó comprobantes para Pagar", vbInformation, "Aviso"
        cmdEmitir.Enabled = True 'PASI20150216
        Exit Sub
    End If
    
    Set oDoc = New DDocumento
    Set oDocPago = New clsDocPago
    Set oNContFunc = New NContFunciones
    Set oCtasIF = New NCajaCtaIF
    Set oDCtaIF = New DCajaCtasIF
    Set oConst = New NConstSistemas
    
    lbBitReten = IIf(oConst.LeeConstSistema(gConstSistBitRetencion6Porcent) = 1, True, False)
    lsCtaEntidadOrig = Trim(lblEntidadPagadoraNombre.Caption)
    lsTpoIf = Mid(txtEntidadPagadoraCod.Text, 1, 2)
    lsCtaBanco = Mid(txtEntidadPagadoraCod.Text, 18, Len(Me.txtEntidadPagadoraCod.Text))
    lsPersCodIF = Mid(txtEntidadPagadoraCod.Text, 4, 13)
    lsEntidadOrig = oDCtaIF.NombreIF(lsPersCodIF)
    lsSubCuentaIF = oCtasIF.SubCuentaIF(lsPersCodIF)
    lsFecha = Format(gdFecSis, "dd/mm/yyyy")
    
    Set oCtasIF = Nothing
    Set oDCtaIF = Nothing
       
    If lsDocTpo = TpoDocCheque Then
        lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, , Mid(gsOpeCod, 3, 1))
        oDocPago.InicioCheque lsDocNRo, True, Mid(txtEntidadPagadoraCod.Text, 4, 13), gsOpeCod, lsPersNombre, gsOpeDesc, gsGlosa, lnImporteB + lnImporteS + lnImporteSEGU - IIf(lmn, lnRetAct, lnRetActME), gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge, , , lsTpoIf, lsPersCodIF, lsCtaBanco
        If oDocPago.vbOk Then
            lsFecha = oDocPago.vdFechaDoc
            lsDocNroTmp = oDocPago.vsNroDoc
            lsDocVoucherTmp = oDocPago.vsNroVoucher
        Else
            cmdEmitir.Enabled = True 'PASI20150216
            Exit Sub
        End If
    ElseIf lsDocTpo = TpoDocCarta Then
        Do While True
            lsPlanillaNro = InputBox("Ingrese el Nro. de Planilla de Pago Proveedores", "Planilla de Pagos", lsPlanillaNro)
            If lsPlanillaNro = "" Then
                cmdEmitir.Enabled = True 'PASI20150216
                Exit Sub
            End If
            lsPlanillaNro = Format(lsPlanillaNro, "00000000")
            If oDoc.GetValidaDocProv("", CLng(lsDocTpo), lsPlanillaNro) Then
                MsgBox "Nro. de carta ya ha sido ingresada, verifique..!", vbInformation, "Aviso"
            Else
                lsDocNroTmp = lsPlanillaNro
                lsDocVoucherTmp = ""
                gnMgIzq = 17
                gnMgDer = 0
                gnMgSup = 12
                Exit Do
            End If
        Loop
    End If
    Set oDoc = Nothing
    
    If MsgBox("¿Esta seguro de realizar los Pagos seleccionados?", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        cmdEmitir.Enabled = True 'PASI20150216
        Exit Sub
    End If
    Set oNCaja = New nCajaGeneral
    Set oImpuesto = New DImpuesto
    Set oPrevio = New clsPrevioFinan
    Set oImp = New NContImprimir
    Set oDis = New NRHProcesosCierre
    ReDim MatPagos(0)
    
    Screen.MousePointer = 11
    cmdEmitir.Enabled = False
    For K = 1 To fg.Rows - 1
        If fg.TextMatrix(K, 2) = "." Then
            lbDocConIGV = False
            lsCabeImpre = " DOCUMENTOS PAGADOS : "
            lsDocVoucher = ""
            lsDocNRo = ""
            lbEfectivo = False
            lsCadBol = ""
            lsPersCod = ""
            lnImporteB = 0: lnImporteS = 0: lnImporteAjusteDolaresS = 0: lnImporteAjusteDolaresB = 0
            lnImporteRH = 0: lnImporteAjusteDolaresRH = 0: lnImporteRHJ = 0: lnImporteAjusteDolaresRHJ = 0
            lnImporteSEGU = 0: lnImporteAjusteDolaresSEGU = 0: lnImportePAGOVARIOS = 0: lnImporteAjusteDolaresPagoVarios = 0
            lsCtaContLeasing = ""
            lnMovNroProv = 0
            
            lsGlosa = fg.TextMatrix(K, 6) 'PASIERS1242014
            lsPersCod = fg.TextMatrix(K, 8)
            lsPersNombre = fg.TextMatrix(K, 5)
            lsCuentaAho = Trim(fg.TextMatrix(K, 29))
            lnMovNroProv = CLng(fg.TextMatrix(K, 10))
            lsCtaContDebeB = oOpe.EmiteOpeCta(gsOpeCod, "D", "0") 'NAGL 20171229
             
            lsCabeImpre = lsCabeImpre & oImpresora.gPrnCondensadaON & fg.TextMatrix(K, 3) & space(5) & oImpresora.gPrnCondensadaOFF
            lsCabeImpre = lsCabeImpre & oImpresora.gPrnSaltoLinea & space(22)

            If lsCtaContDebeB = fg.TextMatrix(K, 13) And Not (lsCtaContDebeBLeasingMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingME = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSME = fg.TextMatrix(K, 13)) Then
                lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
            End If
            'VAPA201724
             If lsCtaContHaberFepMacS = fg.TextMatrix(K, 13) And Not (lsCtaContDebeBLeasingMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingME = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSME = fg.TextMatrix(K, 13)) Then
                lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
                lsCtaContDebeB = lsCtaContHaberFepMacS 'NAGL 20171229
            End If
            'END
            
            '******************NAGL según INC1712260008***********************************
            Set rs = oDCaja.GetProveedoresCtasAperturadasSBSbyItem(gsOpeCod, "byItem")
            If Not rs.BOF And Not rs.EOF Then
                Do While Not rs.EOF
                   lsCtaContProvNewApert = rs!cCtaComp
                   If lsCtaContProvNewApert = fg.TextMatrix(K, 13) And Not (lsCtaContDebeBLeasingMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingME = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSME = fg.TextMatrix(K, 13)) Then
                        lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
                        lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
                        lsCtaContDebeB = lsCtaContProvNewApert
                   End If
                   rs.MoveNext
                Loop
            End If
            Set rs = Nothing
            '*****************END NAGL 20171228*******************************************
            
            If lsCtaContDebeS = fg.TextMatrix(K, 13) Then
                lnImporteS = lnImporteS + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresS = lnImporteAjusteDolaresS + CCur(fg.TextMatrix(K, 20))
            End If
            If lsCtaContDebeRH = fg.TextMatrix(K, 13) Then
                lnImporteRH = lnImporteRH + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresRH = lnImporteAjusteDolaresRH + CCur(fg.TextMatrix(K, 20))
            End If
            If lsCtaContDebeRHJ = fg.TextMatrix(K, 13) Then
                lnImporteRHJ = lnImporteRHJ + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresRHJ = lnImporteAjusteDolaresRHJ + CCur(fg.TextMatrix(K, 20))
            End If
            If lsCtaContDebeSegu = fg.TextMatrix(K, 13) Then
                lnImporteSEGU = lnImporteSEGU + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresSEGU = lnImporteAjusteDolaresSEGU + CCur(fg.TextMatrix(K, 20))
            End If
            If lsCtaContDebePagoVarios = fg.TextMatrix(K, 13) Then
                lnImportePAGOVARIOS = lnImportePAGOVARIOS + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresPagoVarios = lnImporteAjusteDolaresPagoVarios + CCur(fg.TextMatrix(K, 20))
            End If
            If lsCtaContDebeBLeasingMN = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSMN = fg.TextMatrix(K, 13) Then
                lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
                lsCtaContLeasing = fg.TextMatrix(K, 13)
            End If
            If lsCtaContDebeBLeasingME = fg.TextMatrix(K, 13) Or lsCtaContDebeBLeasingSME = fg.TextMatrix(K, 13) Then
                lnImporteB = lnImporteB + CCur(fg.TextMatrix(K, 18))
                lnImporteAjusteDolaresB = lnImporteAjusteDolaresB + CCur(fg.TextMatrix(K, 20))
                lsCtaContLeasing = fg.TextMatrix(K, 13)
            End If
            If lbBitReten And Not lbDocConIGV Then
                If Val(fg.TextMatrix(K, 11)) = TpoDocFactura Or Val(fg.TextMatrix(K, 11)) = TpoDocNotaCredito Or Val(fg.TextMatrix(K, 11)) = TpoDocNotaDebito Then
                    lbDocConIGV = True
                End If
            End If
            
            bAceptaRetencion = True
            If lbBitReten Then
                lbBCAR = VerifBCAR(lsPersCod)
                lsCtaReten = oConst.LeeConstSistema(gConstSistCtaRetencion6Porcent)
                lnTasaImp = oImpuesto.CargaImpuesto(lsCtaReten)!nImpTasa
                lnIngresos = oNCaja.GetMontoIngresoRetencion(lsPersCod, Left(Format(gdFecSis, gsFormatoMovFecha), 6), True)
                lnRetencion = oNCaja.GetMontoIngresoRetencion(lsPersCod, Left(Format(gdFecSis, gsFormatoMovFecha), 6), False)
                lnTopeRetencion = oConst.LeeConstSistema(gConstSistTopeRetencion6Porcent)
            
                If Not lbBCAR Then
                    If lmn Then
                        lnRetAct = (lnImporteB + lnImporteS) + lnIngresos
                    Else
                        lnRetAct = Round((lnImporteB + lnImporteS) * gnTipCambioPonderado, 2) + lnIngresos
                    End If
                    If lnRetAct <= lnTopeRetencion Then
                        lnRetAct = 0
                    Else
                        lnRetAct = Round(lnRetAct * (lnTasaImp / 100), 2) - lnRetencion
                        If lmn Then
                            If lnRetAct > (lnImporteB + lnImporteS) Then
                                lnRetAct = (lnImporteB + lnImporteS)
                            End If
                        Else
                            If lnRetAct > (lnImporteB + lnImporteS) * gnTipCambioPonderado Then
                                lnRetAct = (lnImporteB + lnImporteS) * gnTipCambioPonderado
                            End If
                        End If
                    End If
                Else
                    lnRetAct = 0
                End If
            
                If lmn Then
                    lnRetActME = 0
                Else
                    lnRetActME = Round(lnRetAct / gnTipCambioPonderado, 2)
                End If
                
                If lnRetAct > 0 Then
                    Dim sTexto As String
                    Do While True
                        sTexto = InputBox("El proveeedor: " & lsPersNombre & " esta afecto a una retención de : ", "Retención a Pago", Round(lnRetAct, 2))
                        If sTexto = "" Then
                            Exit Do
                        End If
                        If IsNumeric(sTexto) Then
                            lnRetAct = CCur(sTexto)
                            Exit Do
                        Else
                            MsgBox "Debe ingresar dato Númerico", vbInformation, "¡Aviso!"
                        End If
                    Loop
                    lnRetActME = Round(lnRetAct / gnTipCambioPonderado, 2)
                End If
                If lnRetAct > 0 Then
                    If MsgBox("El proveeedor: " & lsPersNombre & " esta afecto a una retención por el comprobante " & fg.TextMatrix(K, 3) & " de (" & gcMN & ") : " & Format(lnRetAct, "#,##0.00") & vbNewLine & "Desea incluirlo en este Pago?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
                        bAceptaRetencion = False
                    End If
                End If
            Else
                lnRetAct = 0
            End If
            
            If bAceptaRetencion Then
                If lsDocTpo = TpoDocCarta Then
                    lsDocNRo = lsDocNroTmp
                    lsDocVoucher = lsDocVoucherTmp
                ElseIf lsDocTpo = TpoDocCheque Then
                    lsDocNRo = lsDocNroTmp
                    lsDocVoucher = lsDocVoucherTmp
                ElseIf lsDocTpo = TpoDocNotaAbono Then
                    lsDocNRo = oNContFunc.GeneraDocNro(lsDocTpo, , , , True)
                    lsFecha = Format(gdFecSis, "dd/mm/yyyy")
                    lnITF = 0
                    lsCadBol = oDis.ImprimeBoletaCad(CDate(lsFecha), "ABONO CAJA GENERAL", "Depósito CAJA GENERAL*Nro." & lsDocNRo, "", lnImporteB + lnImporteS + lnImporteSEGU - IIf(lmn, lnRetAct, lnRetActME), lsPersNombre, lsCuentaAho, "", 0, 0, "Nota Abono", 0, 0, False, False, , , , True, , , , False, gsNomAge)
                End If
                    
                Dim oCtSaldo As New DCtaSaldo
                Dim oRS As New ADODB.Recordset
                Set oRS = oCtSaldo.ObtenerOperacionesSAF_NEW(CLng(fg.TextMatrix(K, 10)))
                lsCtaCore = ""
                lsCtaSAF = ""
                If Not oRS.EOF Then
                    lsCtaCore = oRS!cCtaCod
                    lsCtaSAF = oRS!cCtaSaf
                End If
                
                Sleep 1000
                lsMovNro = oNContFunc.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
                lsDocTpoTmp = lsDocTpo
                
                lOk = oNCaja.GrabaPagoProveedor_NEW(lsMovNro, lsOpeCod, lsGlosa, lsCtaContDebeB, lsCtaContDebeS, _
                        lsCtaContHaber, lnImporteB, lnImporteS, lsPersCod, lsTpoIf, lsPersCodIF, lsCtaBanco, _
                        rsBilletaje, lsDocTpoTmp, lsDocNRo, Format(CDate(lsFecha), gsFormatoFecha), lsDocVoucher, fg.GetRsNew, lsCuentaAho, lsCtaDiferencia, lnMontoDif, gbBitCentral, True, IIf(Mid(gsOpeCod, 3, 1) = 1, lnRetAct, lnRetActME), lsCtaITFD, lsCtaITFH, gnImpITF, False, IIf(chkAfectoITF.value = 1, True, False), lnITF, lnImporteRH, lsCtaContDebeRH, lsCtaContDebeRHJ, lnImporteRHJ, lsCtaContDebeSegu, lnImporteSEGU, lsCtaContDebePagoVarios, lnImportePAGOVARIOS, lnTipoPago, lsCtaSAF, lsCtaCore, lsCtaContLeasing, K, lnMovNroProv, fsCtaContPenalidadH) 'PASIERS1242014 se Cambio 'txtObservaciones.Text' por 'lsGlosa'
                If lOk = 1 Then
                    ReDim Preserve MatPagos(UBound(MatPagos) + 1)
                    MatPagos(UBound(MatPagos)).lnItem = K
                    MatPagos(UBound(MatPagos)).lsMovNro = lsMovNro
                    MatPagos(UBound(MatPagos)).lsVoucherPago = lsCadBol
                End If
            End If
        End If
    Next
    cmdEmitir.Enabled = True
    Screen.MousePointer = 0
    
    lsImpre = ""
    lnNroPagos = UBound(MatPagos)
    If lnNroPagos > 0 Then
        For K = 1 To lnNroPagos
            lsImpre = lsImpre & oImp.ImprimeAsientoContable(MatPagos(K).lsMovNro, gnLinPage, gnColPage, "PAGO A PROVEEDORES") & oImpresora.gPrnSaltoPagina
        Next
        EnviaPrevio lsImpre, "PAGO A PROVEEDORES", gnLinPage, False
        If lsDocTpo = TpoDocNotaAbono Then
            lsImpre = ""
            For K = 1 To lnNroPagos
                lsImpre = lsImpre & MatPagos(K).lsVoucherPago
            Next
            EnviaPrevio lsImpre & oImpresora.gPrnSaltoPagina, "PAGO A PROVEEDORES", gnLinPage, False
        End If
        If MsgBox("Pago realizado, " & Format(lnNroPagos, "00") & " comprobantes cancelados" & vbNewLine & "¿Desea exportar el detalle a Excel?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            ExportarExcel MatPagos
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Realizo Pago "
                Set objPista = Nothing
                '****
        End If
        cmdCancelar_Click
        cmdProcesar_Click
    End If

    Set oNContFunc = Nothing
    Set oOpe = Nothing
    Set oDocPago = Nothing
    Set oNCaja = Nothing
    Set oConst = Nothing
    Set oImpuesto = Nothing
    Set oPrevio = Nothing
    Set rsBilletaje = Nothing
    Set oDis = Nothing
    Set oImp = Nothing
    Exit Sub
NoGrabo:
    cmdCancelar_Click 'CTI2 ADD 20190430
    cmdEmitir.Enabled = True
    Screen.MousePointer = 0
    MsgBox TextErr(Err.Description), vbInformation, "Aviso"
End Sub
Private Sub cmdCancelar_Click()
    fraBusqueda.Enabled = True
    chkTodos.value = 0
    FormateaFlex fg
    chkAfectoITF.value = 1
    txtObservaciones.Text = ""
    txtEntidadPagadoraCod.Text = ""
    lblEntidadPagadoraNombre.Caption = ""
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub fg_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    Dim i As Integer
    Dim bCheckeado As Boolean
    Dim lsCheck As String
    For i = 1 To fg.Rows - 1
        If fg.TextMatrix(i, 2) = "." Then
            bCheckeado = True
            Exit For
        End If
    Next
    If Not bCheckeado Then
        chkTodos.value = 0
    End If
    If fnTpoPago = gPagoCheque Then
        If fg.TextMatrix(pnRow, 2) = "." Then
            For i = 1 To fg.Rows - 1
                If pnRow <> i Then
                    If fg.TextMatrix(pnRow, 8) <> fg.TextMatrix(i, 8) And fg.TextMatrix(i, 2) = "." Then 'Diferente al Proveedor y que este checkeado se deschekeará
                        fg.TextMatrix(i, 2) = ""
                    End If
                End If
            Next
        End If
    End If
End Sub
Private Sub fg_OnRowChange(pnRow As Long, pnCol As Long)
    txtObservaciones.Text = fg.TextMatrix(pnRow, 6)
End Sub
Private Sub fg_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(fg.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    CargaControles
    limpiarControles
    CargaVariables
End Sub
Private Sub CargaControles()
    Dim oLog As New DLogGeneral
    Dim oConst As New DConstante
    Dim rs As New ADODB.Recordset
    Dim oOpe As New DOperacion
    'Carga Meses
    cboMes.Clear
    cboMes.AddItem "ENERO" & space(150) & "1"
    cboMes.AddItem "FEBRERO" & space(150) & "2"
    cboMes.AddItem "MARZO" & space(150) & "3"
    cboMes.AddItem "ABRIL" & space(150) & "4"
    cboMes.AddItem "MAYO" & space(150) & "5"
    cboMes.AddItem "JUNIO" & space(150) & "6"
    cboMes.AddItem "JULIO" & space(150) & "7"
    cboMes.AddItem "AGOSTO" & space(150) & "8"
    cboMes.AddItem "SEPTIEMBRE" & space(150) & "9"
    cboMes.AddItem "OCTUBRE" & space(150) & "10"
    cboMes.AddItem "NOVIEMBRE" & space(150) & "11"
    cboMes.AddItem "DICIEMBRE" & space(150) & "12"
    'Carga Tipo de Pagos
    Set rs = oConst.RecuperaConstantes(10030)
    cboTpoPago.Clear
    CargaCombo cboTpoPago, rs
    'Carga Entidad Receptora
    txtEntidadReceptoraCod.psRaiz = "Cuentas de Instituciones Financieras"
    txtEntidadReceptoraCod.rs = oLog.ListaIFisxPagoProveedor
    'Carga Entidad Pagadora
    txtEntidadPagadoraCod.psRaiz = "Cuentas de Instituciones Financieras"
    txtEntidadPagadoraCod.rs = oOpe.GetRsOpeObj(gsOpeCod, "1")
    txtObservaciones.Enabled = False 'PASIERS1242014
    
    ReDim fMatProveedor(1, 0)
    
    Set rs = Nothing
    Set oOpe = Nothing
    Set oConst = Nothing
    Set oLog = Nothing
End Sub
Private Sub CargaVariables()
    Dim oOpe As New DOperacion
    Dim oNConst As New NConstSistemas
    Dim rs As New ADODB.Recordset
    Dim lsCtaContPenalidad As String
    
    lsCtaITFD = oOpe.EmiteOpeCta(gsOpeCod, "D", 2)
    lsCtaITFH = oOpe.EmiteOpeCta(gsOpeCod, "H", 2)
    lsCtaContDebeB = oOpe.EmiteOpeCta(gsOpeCod, "D", "0")
    lsCtaContDebeS = oOpe.EmiteOpeCta(gsOpeCod, "D", "1")
    lsCtaContDebeRH = oOpe.EmiteOpeCta(gsOpeCod, "D", "3")
    lsCtaContDebeRHJ = oOpe.EmiteOpeCta(gsOpeCod, "D", "4")
    lsCtaContDebeSegu = oOpe.EmiteOpeCta(gsOpeCod, "D", "5")
    lsCtaContDebePagoVarios = oOpe.EmiteOpeCta(gsOpeCod, "D", "6")
    lsCtaContDebeBLeasingMN = oOpe.EmiteOpeCtaLeasing("421110", 1)
    lsCtaContDebeBLeasingSMN = oOpe.EmiteOpeCtaLeasing("421110", 2)
    lsCtaContDebeBLeasingME = oOpe.EmiteOpeCtaLeasing("422110", 1)
    lsCtaContDebeBLeasingSME = oOpe.EmiteOpeCtaLeasing("422110", 2)
    lsCtaContHaberFepMacS = oOpe.EmiteOpeCta(gsOpeCod, "H", 8) 'VAPA20170724
    lsTipoB = "T"
    lnTipoPago = 0
    fsPersCodCMACMaynas = "1090100012521"
    fsPersCodBCP = "1090100824640"
    lsDocTpo = "-1"
    lnTipoPago = 0
    lmn = IIf(Mid(gsOpeCod, 3, 1) = Moneda.gMonedaExtranjera, False, True)
    lsFileCarta = App.path & "\" & gsDirPlantillas & gsOpeCod & ".TXT"
    lsCtaContPenalidad = oNConst.LeeConstSistema(452)
    fsCtaContPenalidadH = IIf(lsCtaContPenalidad <> "", Left(lsCtaContPenalidad, 2) & Mid(gsOpeCod, 3, 1) & Mid(lsCtaContPenalidad, 4, Len(lsCtaContPenalidad)), "")
    
    Set rs = oOpe.CargaOpeDoc(gsOpeCod, , OpeDocMetDigitado)
    lsDocs = RSMuestraLista(rs, 1)
    
    Set rs = Nothing
    Set oNConst = Nothing
    Set oOpe = Nothing
End Sub
Private Sub limpiarControles()
    Dim lnMes As Integer
    lnMes = Month(gdFecSis)
    cboMes.ListIndex = IndiceListaCombo(cboMes, lnMes)
    txtAnio.Text = Year(gdFecSis)
    cboTpoPago.ListIndex = -1
    txtEntidadReceptoraCod.Text = ""
    chkTodos.value = 0
    FormateaFlex fg
    chkAfectoITF.value = 1
    txtObservaciones.Text = ""
    txtEntidadPagadoraCod.Text = ""
    lblEntidadPagadoraNombre.Caption = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("¿Desea salir del Proceso de Pago a Proveedores?", vbYesNo + vbQuestion, "Aviso") = vbNo Then
        Cancel = 1
    End If
End Sub
Private Sub txtAnio_Change()
    If Len(txtAnio.Text) = 4 Then
        If cboTpoPago.Visible And cboTpoPago.Enabled Then
            cboTpoPago.SetFocus
        End If
    End If
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = vbKeyBack Then
        If Len(txtAnio.Text) = 0 Then
            cboMes.SetFocus
        End If
    End If
    If KeyAscii = 13 Then
        If cboTpoPago.Visible And cboTpoPago.Enabled Then
            cboTpoPago.SetFocus
        End If
    End If
End Sub
Private Sub txtAnio_LostFocus()
    txtAnio.Text = Format(txtAnio.Text, "0000")
End Sub
Private Sub txtEntidadReceptoraCod_EmiteDatos()
    lblEntidadReceptoraNombre.Caption = ""
    If txtEntidadReceptoraCod.Text <> "" Then
        lblEntidadReceptoraNombre.Caption = txtEntidadReceptoraCod.psDescripcion
    End If
End Sub
Private Sub txtEntidadReceptoraCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdProcesar.SetFocus
    End If
End Sub
Private Sub txtEntidadPagadoraCod_EmiteDatos()
    lblEntidadPagadoraNombre.Caption = ""
    If txtEntidadPagadoraCod.Text <> "" Then
        lblEntidadPagadoraNombre.Caption = txtEntidadPagadoraCod.psDescripcion
    End If
    
    Dim oCtaIf As NCajaCtaIF
    Set oCtaIf = New NCajaCtaIF
    lblEntidadPagadoraNombre.Caption = oCtaIf.EmiteTipoCuentaIF(Mid(txtEntidadPagadoraCod.Text, 18, 10)) + " " + txtEntidadPagadoraCod.psDescripcion
    Set oCtaIf = Nothing
End Sub
Private Sub txtEntidadPagadoraCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdEmitir.SetFocus
    End If
End Sub
Private Function validaProcesar() As Boolean
    validaProcesar = True
    If cboMes.ListIndex = -1 Then
        validaProcesar = False
        MsgBox "Ud. debe seleccionar el Mes de Búsqueda", vbInformation, "Aviso"
        cboMes.SetFocus
        Exit Function
    End If
    If Val(txtAnio.Text) <= 1900 Then
        validaProcesar = False
        MsgBox "Ud. debe de ingresar el año de Búsqueda", vbInformation, "Aviso"
        txtAnio.SetFocus
        Exit Function
    End If
    If cboTpoPago.ListIndex = -1 Then
        validaProcesar = False
        MsgBox "Ud. debe de seleccionar el Tipo de Pago", vbInformation, "Aviso"
        cboTpoPago.SetFocus
        Exit Function
    End If
    If fraEntidadReceptora.Visible Then
        If Len(Trim(txtEntidadReceptoraCod.Text)) = 0 Then
            validaProcesar = False
            MsgBox "Ud. debe de seleccionar la Institución Financiera de Búsqueda", vbInformation, "Aviso"
            If txtEntidadReceptoraCod.Visible And txtEntidadReceptoraCod.Enabled Then txtEntidadReceptoraCod.SetFocus
            Exit Function
        End If
    End If
End Function
Private Function ValidaInterfaz() As Boolean
    ValidaInterfaz = True
    If Len(Trim(txtObservaciones.Text)) = 0 Then
        ValidaInterfaz = False
        MsgBox "Ud. debe indicar las observaciones respectivas", vbInformation, "Aviso"
        txtObservaciones.SetFocus
        Exit Function
    End If
    If fraEntidadPagadora.Visible Then
        If Len(Trim(txtEntidadPagadoraCod.Text)) = 0 Then
            ValidaInterfaz = False
            MsgBox "Ud. debe seleccionar la cuenta de la Institución Financiera", vbInformation, "Aviso"
            If txtEntidadPagadoraCod.Visible And txtEntidadPagadoraCod.Enabled Then txtEntidadPagadoraCod.SetFocus
            Exit Function
        End If
    End If
End Function
Private Sub ExportarExcel(ByRef pPagos() As PagoProv)
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim lnFila As Long, lnColumna As Long, lnColumnaMax As Long
    Dim i As Long, j As Long, K As Long
    Dim lsArchivo As String
    Dim bOK As Boolean
    Dim lnNroPagos As Integer

    lnNroPagos = UBound(pPagos)
    If lnNroPagos > 0 Then
        Set xlsAplicacion = New Excel.Application
        lsArchivo = "\spooler\RptPagoComprobantes" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xls"
        Set xlsLibro = xlsAplicacion.Workbooks.Add
    
        Set xlsHoja = xlsLibro.Worksheets.Add
        xlsHoja.Name = "Pago de Comprobantes"
        xlsHoja.Cells.Font.Name = "Arial"
        xlsHoja.Cells.Font.Size = 9
        
        lnFila = 2
        
        For i = 0 To fg.Rows - 1
            lnColumna = 2
            For j = 3 To fg.Cols - 1
                If fg.ColWidth(j) > 0 Then
                    bOK = False
                    For K = 1 To lnNroPagos
                        If pPagos(K).lnItem = i Then
                            bOK = True
                            Exit For
                        End If
                    Next
                    If bOK = True Or i = 0 Then
                        xlsHoja.Cells(lnFila, lnColumna) = IIf(i = 0, "'", "") & fg.TextMatrix(i, j)
                        lnColumna = lnColumna + 1
                        lnColumnaMax = lnColumna
                    End If
                End If
            Next
            If bOK = True Or i = 0 Then
                lnFila = lnFila + 1
            End If
        Next
        'lnFila = 2 + lnNroPagos
        
        xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).Interior.Color = RGB(191, 191, 191)
        xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).Font.Bold = True
        xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).HorizontalAlignment = xlCenter
        xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(lnFila - 1, lnColumnaMax - 1)).Borders.Weight = xlThin
        xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(lnFila - 1, lnColumnaMax - 1)).EntireColumn.AutoFit
        
        xlsHoja.SaveAs App.path & lsArchivo
        xlsAplicacion.Visible = True
        xlsAplicacion.Windows(1).Visible = True
        
        Set xlsAplicacion = Nothing
        Set xlsLibro = Nothing
        Set xlsHoja = Nothing
        Exit Sub
    End If
End Sub
