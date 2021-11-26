VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInversiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Inversiones"
   ClientHeight    =   7260
   ClientLeft      =   3120
   ClientTop       =   2190
   ClientWidth     =   9885
   Icon            =   "frmInversiones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   9885
   Begin VB.Frame Frame3 
      Caption         =   "Cuentas a Aperturar"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   9675
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   1290
         TabIndex        =   18
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   150
         TabIndex        =   17
         Top             =   1680
         Width           =   1095
      End
      Begin Sicmact.FlexEdit fgCta 
         Height          =   1395
         Left            =   150
         TabIndex        =   19
         Top             =   240
         Width           =   9375
         _ExtentX        =   15055
         _ExtentY        =   2461
         Cols0           =   10
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Codigo-Cuenta_Nro-Importe-Plazo-Fec_Apertura-Fec_Venc-Tasa-Int_Pactado-ultimafila"
         EncabezadosAnchos=   "300-1200-2700-1400-600-1300-1300-1000-1200-0"
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
         ColumnasAEditar =   "X-X-2-3-4-5-6-7-8-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-2-2-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R-R-L-L-R-R-L"
         FormatosEdit    =   "0-0-0-2-3-0-0-2-2-0"
         CantEntero      =   12
         CantDecimales   =   6
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
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
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   9675
      Begin Sicmact.TxtBuscar txtBuscaEntidad 
         Height          =   360
         Left            =   1080
         TabIndex        =   12
         Top             =   300
         Width           =   2580
         _ExtentX        =   4551
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
      Begin VB.Label lblDesCtaIfTransf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1095
         TabIndex        =   15
         Top             =   720
         Width           =   8370
      End
      Begin VB.Label lblDescIfTransf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3735
         TabIndex        =   14
         Top             =   300
         Width           =   5730
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta N° :"
         Height          =   210
         Left            =   180
         TabIndex        =   13
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Glosa"
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
      Height          =   1380
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Width           =   9675
      Begin VB.TextBox txtMovDesc 
         Height          =   960
         Left            =   135
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   255
         Width           =   9210
      End
   End
   Begin VB.Frame Frame1 
      Height          =   680
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Width           =   9675
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1440
         TabIndex        =   20
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Guardar"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   8280
         TabIndex        =   4
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.Frame fradatosGen 
      Caption         =   "Datos Generales"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9705
      Begin VB.Frame Frame5 
         Caption         =   "Fecha Movimiento"
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
         Left            =   7500
         TabIndex        =   24
         Top             =   600
         Width           =   1935
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   330
            Left            =   360
            TabIndex        =   25
            Top             =   240
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Condicion"
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
         Left            =   3360
         TabIndex        =   21
         Top             =   600
         Width           =   2055
         Begin VB.OptionButton optCondicion 
            Caption         =   "Fija"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optCondicion 
            Caption         =   "Variable"
            Height          =   375
            Index           =   1
            Left            =   840
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbTipoInversion 
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   1995
      End
      Begin Sicmact.TxtBuscar txtBuscaIF 
         Height          =   360
         Left            =   1245
         TabIndex        =   1
         Top             =   240
         Width           =   1995
         _ExtentX        =   3519
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
      Begin VB.Label lblDescIF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3270
         TabIndex        =   10
         Top             =   255
         Width           =   6165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Administradora"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   323
         Width           =   1035
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Inversion:"
         Height          =   270
         Left            =   90
         TabIndex        =   5
         Top             =   682
         Width           =   1170
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmInversiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsSubCtaIF As String
Dim lnTipoCtaIf As Integer
Dim oOpe As DOperacion
Dim oCtaIf As NCajaCtaIF
'Dim lsCuentaCod As String

Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset

Dim lbCreaSubCta As Boolean
Dim lnOptionGrabar As Integer '0=Registrar , 1=Modificar
Dim lnMovNro As Long
Dim lsMovNroAnt As String
Dim lsOpeCod As String
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub ConfigCabeceraFG()
   With Me.fgCta
'    .FormaCabecera
    .FormateaColumnas
    'Si es Fondos Mutuos=3
    If Trim(Right(Me.cmbTipoInversion.Text, 2)) = "3" Then
        
        .EncabezadosNombres = "#-Codigo-Cuenta_Nro-Importe-Val_Cuota_1-Nro_Cuotas-Fec_Apertura-ultimafila1-ultimafila2-ultimafila3"
        .ColumnasAEditar = "X-X-2-3-4-5-6-X-X-X"
        .EncabezadosAnchos = "300-1200-2700-1400-1400-1400-1300-0-0-0-0"
        .EncabezadosAlineacion = "C-L-L-R-R-R-L-L-L-L"
        .FormatosEdit = "0-0-0-2-0-0-0-0-0-0"
        .ListaControles = "0-0-0-0-0-0-2-0-0-0"
        
    Else
        .EncabezadosNombres = "#-Codigo-Cuenta_Nro-Importe-Plazo-Fec_Apertura-Fec_Venc-Tasa-Int_Pactado-ultimafila"
        .ColumnasAEditar = "X-X-2-3-4-5-6-7-8-X"
        .EncabezadosAnchos = "300-1200-2700-1400-600-1300-1300-1000-1200-0"
        .EncabezadosAlineacion = "C-L-L-R-R-L-L-R-R-L"
        .FormatosEdit = "0-0-0-2-3-0-0-2-2-0"
        .ListaControles = "0-0-0-0-0-2-2-0-0-0"
    End If
'    .FormateaColumnas
   End With
End Sub

Private Sub cmbTipoInversion_Click()
    fgCta.Clear
    fgCta.Rows = 2
    fgCta.FormaCabecera
    ConfigCabeceraFG
End Sub

Private Sub cmdAceptar_Click()
    Dim lsCuentaAho As String
    
    Dim lsMovNro As String
    Dim oCon     As NContFunciones
    Dim oCaja As nCajaGeneral
    Dim rsAdeud  As ADODB.Recordset
    On Error GoTo ErrApertura
        
    If Valida = False Then
       Exit Sub
    End If
    
    If Not validaFechaMov Then
        MsgBox "La Fecha del Movimiento debe estar dentro del Mes Vigente", vbInformation, "AVISO"
        Exit Sub
    End If
    
    Set oCon = New NContFunciones
    Set oCaja = New nCajaGeneral
    
   'LARI20200309 validamos si registró interes pactados
   With Me.fgCta
    If Trim(Right(cmbTipoInversion.Text, 2)) <> "3" Then 'Si es fondo Mutuo
        Dim fila As Integer
        Dim I As Integer
        Dim valor1 As String
        Dim Nulos As Integer
        Dim PasaXUser As Boolean
        
        PasaXUser = True
        
        fila = .Rows - 1
        For I = 1 To fila
            If (.TextMatrix(fila, 3) <> "" And .TextMatrix(fila, 4) <> "" And .TextMatrix(fila, 7) <> "") And CDbl(.TextMatrix(fila, 8)) = 0 Then
                Nulos = Nulos + 1
            End If
        Next I
        
        If Nulos > 0 Then
            If MsgBox("Interes pactado = 0.00" & Chr(13) & "¿Desea desea continuar con el registro? ", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                PasaXUser = False
            End If
        End If
    End If
   End With
   
    'guardando datos
    If PasaXUser = True Then 'LARI 20200309 SI ACEPTO GUARDAR LOS DATOS CON EL INTERES PACTADO EN CERO
        If MsgBox(" ¿ Desea Grabar Operación de Apertura ? ", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        With Me.fgCta
            If lnOptionGrabar = 0 Then 'Nuevo
                'lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                lsMovNro = oCon.GeneraMovNro(txtFecha.Text, gsCodAge, gsCodUser)
                
                oCaja.guardarInversionApertura lsMovNro, gsOpeCod, txtMovDesc, _
                        fgCta.GetRsNew, lbCreaSubCta, lsSubCtaIF, Mid(txtBuscaIF, 4, 13), _
                        Mid(txtBuscaIF, 1, 2), CInt(Mid(gsOpeCod, 3, 1)), txtFecha.Text, txtBuscaEntidad, _
                        CInt(Trim(Right(Me.cmbTipoInversion.Text, 2))), IIf(Me.optCondicion(0).value = True, 0, 1)
                        
                MsgBox "Se han Registrado los Datos con Exito", vbInformation, "AVISO"
        
                ImprimeAsientoContable lsMovNro, "", "", "", True, False
                If MsgBox(" ¿ Desea Realizar Otra Operacion ? ", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                    LimparControles
                Else
                     Unload Me
                End If
            ElseIf lnOptionGrabar = 1 Then 'Mantenimiento
                
                oCaja.actualizaInversionDatos lnMovNro, lsOpeCod, Mid(txtBuscaIF, 4, 13), Mid(txtBuscaIF, 1, 2), .TextMatrix(1, 1), _
                                              .TextMatrix(1, 2), IIf(Trim(Right(Me.cmbTipoInversion.Text, 2)) <> "3", .TextMatrix(1, 4), 0), _
                                              IIf(Trim(Right(Me.cmbTipoInversion.Text, 2)) <> "3", .TextMatrix(1, 5), .TextMatrix(1, 6)), .TextMatrix(1, 6), .TextMatrix(1, 3), _
                                              IIf(Trim(Right(Me.cmbTipoInversion.Text, 2)) <> "3", .TextMatrix(1, 7), 0), IIf(Trim(Right(Me.cmbTipoInversion.Text, 2)) <> "3", .TextMatrix(1, 8), 0), _
                                              Me.txtMovDesc.Text, lsMovNroAnt, CInt(Trim(Right(Me.cmbTipoInversion.Text, 2))), IIf(Me.optCondicion(0).value = True, 0, 1), _
                                              IIf(Trim(Right(Me.cmbTipoInversion.Text, 2)) = "3", .TextMatrix(1, 4), 0), IIf(Trim(Right(Me.cmbTipoInversion.Text, 2)) = "3", .TextMatrix(1, 5), 0)
                        
                
                MsgBox "Se han Guardado los Datos con Exito", vbInformation, "AVISO"
                LimparControles
                Unload Me
            
            End If
        End With
                
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grado la Operación "
                Set objPista = Nothing
                '****
       End If
    End If
Exit Sub
ErrApertura:
        MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdAgregar_Click()
    Dim lsCuentaCod As String
    If txtBuscaIF.Text <> "" And lsSubCtaIF <> "" Then
        
        If fgCta.TextMatrix(fgCta.Rows - 1, 1) <> "" Then
           lsCuentaCod = Left(fgCta.TextMatrix(fgCta.Rows - 1, 1), Len(fgCta.TextMatrix(fgCta.Rows - 1, 1)) - 2) + Format(nVal(Right(fgCta.TextMatrix(fgCta.Rows - 1, 1), 2)) + 1, "00")
        Else
            lsCuentaCod = oCtaIf.GetNewCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1), lsSubCtaIF)
        End If
        Me.cmdAgregar.Enabled = False
        fgCta.AdicionaFila
        fgCta.TextMatrix(fgCta.Rows - 1, 1) = lsCuentaCod
        fgCta.TextMatrix(fgCta.Rows - 1, 2) = "Pendiente"
        fgCta.col = 2
        fgCta.SetFocus
    Else
        MsgBox "Aun no selecciono Institución donde Aperturar", vbInformation, "¡Aviso!"
        txtBuscaIF.SetFocus
    End If
End Sub
Private Sub cmdEliminar_Click()
    
    fgCta.EliminaFila fgCta.row
    Me.cmdAgregar.Enabled = True
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fgCta_OnCellChange(pnRow As Long, pnCol As Long)
    With Me.fgCta
        Dim columna As Integer
        columna = .col
        If Trim(Right(cmbTipoInversion.Text, 2)) = "3" Then 'Si es fondo Mutuo
           
            If .col = 4 And .TextMatrix(.row, 3) <> "" And .TextMatrix(.row, 4) <> "" And IsNumeric(.TextMatrix(.row, 4)) Then
                .TextMatrix(.row, 4) = Format(.TextMatrix(.row, 4), "##,##0.000000")
                .TextMatrix(.row, 5) = Format(CDbl(.TextMatrix(.row, 3)) / CDbl(.TextMatrix(.row, 4)), "##,##0.0000")
            ElseIf .col = 4 Then
                .TextMatrix(.row, 4) = ""
            End If
            
            If .col = 5 And .TextMatrix(.row, 5) <> "" And IsNumeric(.TextMatrix(.row, 5)) Then
                .TextMatrix(.row, 5) = Format(.TextMatrix(.row, 5), "##,##0.0000")
            ElseIf .col = 5 Then
                .TextMatrix(.row, 5) = ""
            End If
        
        
        Else 'Si NO es Fondo Mutuo
            If .col = 4 And ValidaFecha(.TextMatrix(.row, 5)) = "" Then
                .TextMatrix(.row, 6) = DateAdd("d", .TextMatrix(.row, 4), .TextMatrix(.row, 5))
            ElseIf .col = 5 And .TextMatrix(.row, 4) <> "" And ValidaFecha(.TextMatrix(.row, 5)) = "" Then
                .TextMatrix(.row, 6) = DateAdd("d", .TextMatrix(.row, 4), .TextMatrix(.row, 5))
            
            'LARI20200309 Cuando se encuentra en la columa del interes pactado, tambien recalcular el valor al salirse de la celda de acuerdo a la tasa ingresadas
            ElseIf (.col = 7 Or .col = 3 Or .col = 4 Or .col = 8) And (.TextMatrix(.row, 3) <> "" And .TextMatrix(.row, 4) <> "" And .TextMatrix(.row, 7) <> "") Then
                .TextMatrix(.row, 8) = Format(getInteresPactado(.TextMatrix(.row, 3), .TextMatrix(.row, 4), .TextMatrix(.row, 7)), "##,##0.00")

            End If
        End If
        
    End With
End Sub
Private Function getInteresPactado(pnImporte As Currency, pnPlazo As Integer, pnTasa As Double) As Currency
'    Dim pnTasa As Currency
'    pnTasa = CCur(pnTasaTemp)
    getInteresPactado = ((1 + (pnTasa / 100) * 1) ^ (pnPlazo * 1 / 360 * 1) - 1) * pnImporte
End Function
Public Function inicio(pnMovNro As Long, psOpecod As String) As Boolean
    lnOptionGrabar = 1
    
    lsOpeCod = psOpecod
    lnMovNro = pnMovNro
    Me.Show 1
    lnMovNro = 0
    lnOptionGrabar = 0
    inicio = True
End Function



Private Sub Form_Load()
    lnTipoCtaIf = gTpoCtaIFCtaPF
    
    cargarTipoInversion
    
    lnTipoCtaIf = 6
    If lnOptionGrabar <> 1 Then
        cargarInstituciones
        lnOptionGrabar = 0
        Me.txtFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
        Me.Caption = "Registro de Inversiones " + IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME")
    Else
        Me.Caption = "Mantenimiento de Inversiones " + IIf(Mid(gsOpeCod, 3, 1) = "1", "MN", "ME")
        CargarDatos
    End If
End Sub
Private Sub CargarDatos()
     Dim oCaja As New nCajaGeneral
     Dim rs As New Recordset
     
     Set rs = oCaja.getInversionDatos(lnMovNro, lsOpeCod)
     If Not (rs.BOF And rs.EOF) Then
        Me.lblDescIF = PstaNombre(rs!cPersNombre)
        Me.txtBuscaIF.Text = rs!cIFTpoAdm + "." + rs!cPersCodAdm
        Me.cmbTipoInversion.ListIndex = rs!nTpoInversion - 1
        Me.optCondicion.Item(rs!nCondicion).value = True
        
        
        With Me.fgCta
            fgCta.AdicionaFila
            fgCta.col = 2
           
            .TextMatrix(1, 1) = rs!cCtaIfCodAdm
            .TextMatrix(1, 2) = rs!cCtaIFDescAdm
            .TextMatrix(1, 3) = Format(rs!nImporte, "##,##0.00")
            If rs!nTpoInversion <> 3 Then
                .TextMatrix(1, 4) = rs!nPlazo
                .TextMatrix(1, 5) = rs!dFecApe
                .TextMatrix(1, 6) = rs!dFecVenc
                .TextMatrix(1, 7) = rs!nTasa
                .TextMatrix(1, 8) = rs!nIntPactado
            Else
                .TextMatrix(1, 4) = Format(rs!nValorCuota, "##,##0.000000")
                .TextMatrix(1, 5) = Format(rs!nNroCuotas, "##,##0.0000")
                .TextMatrix(1, 6) = rs!dFecApe
                
            End If
            Me.txtFecha.Text = rs!dFecApe
        End With
        Me.txtBuscaEntidad.Text = rs!cIFTpoTrans + "." + rs!cPersCodTrans + "." + rs!cCtaIfCodTrans
        txtBuscaEntidad_EmiteDatos
        Me.txtMovDesc.Text = rs!cGlosa
        lsMovNroAnt = rs!cMovNro
        
        Me.txtBuscaIF.Enabled = False
        Me.txtBuscaEntidad.Enabled = False
        Me.cmdAgregar.Enabled = False
        Me.cmdEliminar.Enabled = False
     End If
End Sub
Private Sub cargarInstituciones()
    
    Set oOpe = New DOperacion
    Set oCtaIf = New NCajaCtaIF
    
   txtBuscaIF.psRaiz = "Instituciones Financieras"
   txtBuscaEntidad.psRaiz = "Cuentas de Entidades Financieras"
   
   txtBuscaIF.rs = oOpe.GetOpeObj(gsOpeCod, "1")
   txtBuscaEntidad.rs = oOpe.GetOpeObj(gsOpeCod, "2")
   
   
End Sub
Private Sub cargarTipoInversion()
    Dim rsTpoInversion As ADODB.Recordset
    Dim oCons As DConstante
    
    Set rsTpoInversion = New ADODB.Recordset
    Set oCons = New DConstante
    
    Set rsTpoInversion = oCons.CargaConstante(9990)
    If Not (rsTpoInversion.EOF And rsTpoInversion.BOF) Then
        cmbTipoInversion.Clear
        Do While Not rsTpoInversion.EOF
            cmbTipoInversion.AddItem Trim(rsTpoInversion(2)) & space(100) & Trim(rsTpoInversion(1))
            rsTpoInversion.MoveNext
        Loop
        cmbTipoInversion.ListIndex = 0
    End If
End Sub

Private Sub txtBuscaEntidad_EmiteDatos()
    Set oCtaIf = New NCajaCtaIF
    lblDescIfTransf = oCtaIf.NombreIF(Mid(txtBuscaEntidad.Text, 4, 13))
    lblDesCtaIfTransf = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))) & " " & txtBuscaEntidad.psDescripcion
    
End Sub

Private Sub txtBuscaIF_EmiteDatos()
    lblDescIF = txtBuscaIF.psDescripcion
    
    lsSubCtaIF = ""
    lbCreaSubCta = False
    If txtBuscaIF <> "" Then
        lbCreaSubCta = Not oCtaIf.GetVerificaSubCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1))
        lsSubCtaIF = oCtaIf.GetSubCuentaIF(Mid(txtBuscaIF, 4, 13), Val(Mid(txtBuscaIF, 1, 2)), lnTipoCtaIf, Mid(gsOpeCod, 3, 1))

        fgCta.Clear
        fgCta.Rows = 2
        fgCta.FormaCabecera
        Me.cmdAgregar.Enabled = True
        
    End If
End Sub
Function Valida() As Boolean
    Dim K As Integer
    Valida = False
    
    If Len(Trim(txtBuscaIF)) = 0 Then
        MsgBox "Administradora no seleccionada", vbInformation, "Aviso"
        txtBuscaIF.SetFocus
        Exit Function
    End If
    
    If Me.cmbTipoInversion.ListIndex = -1 Then
        MsgBox "Seleccione un Tipo de Inversion", vbInformation, "Aviso"
        cmbTipoInversion.SetFocus
        Exit Function
    End If
    
    If nVal(fgCta.TextMatrix(1, 3)) = 0 Or fgCta.TextMatrix(1, 3) = "" Then
        MsgBox "Ingrese el Importe", vbInformation, "Aviso"
        fgCta.SetFocus
        fgCta.col = 3
        fgCta.row = 1
        Exit Function
    End If
    
    If nVal(fgCta.TextMatrix(1, 4)) = 0 Or fgCta.TextMatrix(1, 4) = "" Then
        MsgBox "Plazo de Cuenta no válido ", vbInformation, "Aviso"
        fgCta.SetFocus
        fgCta.col = 4
        fgCta.row = 1
        Exit Function
    End If
            
    If fgCta.TextMatrix(K, 5) = "" Or fgCta.TextMatrix(K, 5) = "__/__/____" Then
        MsgBox "Fecha de Apertura no Valida ", vbInformation, "Aviso"
        fgCta.SetFocus
        fgCta.col = 5
        fgCta.row = 1
        Exit Function
    End If
    If fgCta.TextMatrix(K, 6) = "" Or fgCta.TextMatrix(K, 6) = "__/__/____" Then
        MsgBox "Fecha de Vencimiento no Valida ", vbInformation, "Aviso"
        fgCta.SetFocus
        fgCta.col = 6
        fgCta.row = 1
        Exit Function
    End If
            
    If fgCta.TextMatrix(K, 7) = "" Or fgCta.TextMatrix(K, 7) = "0.00" Then
        MsgBox "Tasa no Valida ", vbInformation, "Aviso"
        fgCta.SetFocus
        fgCta.col = 7
        fgCta.row = 1
        Exit Function
    End If
            
      
    If Len(Trim(txtMovDesc)) = 0 Then
        MsgBox "Descripción de Operación no válida", vbInformation, "Aviso"
        txtMovDesc.SetFocus
        Exit Function
    End If
    
    If Len(Trim(Me.txtBuscaEntidad)) = 0 Then
        MsgBox "Entidad Financiera no seleccionada", vbInformation, "Aviso"
        txtBuscaIF.SetFocus
        Exit Function
    End If
    Valida = True
End Function
Private Sub cmdCancelar_Click()
    LimparControles
End Sub
Private Sub LimparControles()
        Me.txtBuscaIF = ""
        Me.txtBuscaEntidad = ""
        Me.lblDescIF = ""
        Me.cmbTipoInversion.ListIndex = -1
        Me.optCondicion(0).value = True
                
        fgCta.Clear
        fgCta.Rows = 2
        fgCta.FormaCabecera
        Me.cmdAgregar.Enabled = True
        Me.txtBuscaEntidad = ""
        Me.lblDescIfTransf = ""
        Me.lblDesCtaIfTransf = ""
        Me.txtMovDesc.Text = ""
        
       
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Not validaFechaMov Then
        MsgBox "La Fecha del Movimiento debe estar dentro del Mes Vigente", vbInformation, "AVISO"
        Me.txtFecha.Text = "__/__/____"
       End If
    End If
End Sub
Private Function validaFechaMov() As Boolean
    validaFechaMov = True
    If DateDiff("M", txtFecha.Text, gdFecSis) <> 0 Then
        validaFechaMov = False
        Exit Function
    End If
    
    'If fgCta.TextMatrix(1, 6) <> "" And fgCta.TextMatrix(1, 6) <> "__/__/____" And Left(txtBuscaIF.Text, 2) <> "09" Then
    If fgCta.TextMatrix(1, 5) <> "" And fgCta.TextMatrix(1, 5) <> "__/__/____" And Left(txtBuscaIF.Text, 2) <> "09" Then 'EJVG20130416
        If DateDiff("M", fgCta.TextMatrix(1, 5), gdFecSis) <> 0 Then
            validaFechaMov = False
        End If
        Me.fgCta.row = 1
        Me.fgCta.col = 5
        Me.fgCta.SetFocus
         Exit Function
    End If
    
End Function
