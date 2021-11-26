VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmInversionesConfirmacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   2235
   ClientTop       =   2985
   ClientWidth     =   11685
   Icon            =   "frmInversionesConfirmacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   11685
   Begin VB.Frame FraLista 
      Height          =   5280
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   11490
      Begin VB.CommandButton cmdConfAper 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   8220
         TabIndex        =   12
         Top             =   4800
         Width           =   1470
      End
      Begin VB.TextBox txtMovDesc 
         Height          =   870
         Left            =   1575
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3750
         Width           =   9660
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   9720
         TabIndex        =   9
         Top             =   4770
         Width           =   1470
      End
      Begin Sicmact.TxtBuscar txtCuenta 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   4785
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   661
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
         sTitulo         =   ""
      End
      Begin Sicmact.FlexEdit fgInversiones 
         Height          =   3240
         Left            =   165
         TabIndex        =   11
         Top             =   255
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5715
         Cols0           =   26
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   $"frmInversionesConfirmacion.frx":030A
         EncabezadosAnchos=   "350-900-800-3500-1350-600-900-1000-1200-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-C-C-C-C-R-L-L-L-L-L-L-L-L-L-L-L-L-L-L-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-2-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "N°"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción"
         Height          =   315
         Left            =   210
         TabIndex        =   14
         Top             =   3810
         Width           =   1065
      End
      Begin VB.Label lblCuenta 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
         Height          =   210
         Left            =   210
         TabIndex        =   13
         Top             =   4890
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11490
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
         Left            =   8160
         TabIndex        =   17
         Top             =   120
         Width           =   1815
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   330
            Left            =   360
            TabIndex        =   18
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
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdProcesar 
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
         Height          =   375
         Left            =   10080
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtDesde 
         Height          =   345
         Left            =   2655
         TabIndex        =   2
         Top             =   225
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   4210816
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin MSMask.MaskEdBox txthasta 
         Height          =   345
         Left            =   4365
         TabIndex        =   6
         Top             =   225
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         _Version        =   393216
         ForeColor       =   4210816
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Label Label3 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   1995
         TabIndex        =   5
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   3765
         TabIndex        =   4
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lbltitulo 
         AutoSize        =   -1  'True
         Caption         =   "CONFIRMACION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   270
         Left            =   90
         TabIndex        =   3
         Top             =   255
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmInversionesConfirmacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub cmdConfAper_Click()
    Dim nImporte   As Currency
    Dim sPersCod   As String
    Dim sTipoIF    As String
    Dim sCtaIfCod  As String
    Dim nMovRef    As Double
    
    If Validar() = False Then Exit Sub
    
    
    nImporte = fgInversiones.TextMatrix(fgInversiones.row, 8)
    sPersCod = fgInversiones.TextMatrix(fgInversiones.row, 14)
    sTipoIF = fgInversiones.TextMatrix(fgInversiones.row, 13)
    sCtaIfCod = fgInversiones.TextMatrix(fgInversiones.row, 11)
    nMovRef = fgInversiones.TextMatrix(fgInversiones.row, 9)
    
    
    ConfirmarInversion nImporte, sPersCod, sTipoIF, sCtaIfCod, nMovRef
    
                'ARLO20170217
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
                Set objPista = Nothing
                '****
    
    Exit Sub
ConfirmaAperErr:
        MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub
Private Sub ConfirmarInversion(pnImporte As Currency, psPersCod As String, psTipoIF As String, psCtaIFCod As String, pnMovRef As Double)
    Dim sCtaHaber  As String
    Dim lsSubCta    As String
    Dim lsMovNro    As String
    Dim lsMovNroACAdd    As String 'PASI20150921 ERS0472015
    Dim oCont       As NContFunciones
    Dim oOpe        As DOperacion
    Dim lnTasaVac   As Double
    Dim lsMonedaPago As String
    'On Error GoTo ConfirmaAperErr
    
    Set oOpe = New DOperacion
    Set oCont = New NContFunciones

    sCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
    
    If sCtaHaber = "" Then
        MsgBox "Cuenta de Pendiente no esta definida. Consultar con Sistemas", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    
    If Not oOpe.ValidaCtaCont(Me.txtCuenta.Text) Then
        MsgBox "No se Encuentra o esta Inactiva la Cuenta Contable: " + Me.txtCuenta.Text, vbInformation, "¡Aviso!"
        Exit Sub
    End If
   
    Dim oCta As DCtaCont
    Dim oCaja As nCajaGeneral
    Set oCta = New DCtaCont
    Set oCaja = New nCajaGeneral
    oCta.ExisteCuenta txtCuenta, True
    If MsgBox("Desea Confirmar la Apertura de la Cuenta Seleccionada?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        With Me.fgInversiones
            'lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            lsMovNro = oCont.GeneraMovNro(txtFecha.Text, gsCodAge, gsCodUser)
            If .TextMatrix(.row, 15) = 6 Then 'PASI20150921 ERS0712015
                Sleep (1000)
                lsMovNroACAdd = oCont.GeneraMovNro(txtFecha.Text, gsCodAge, gsCodUser)
            End If
            If oCaja.GrabaConfInversion(lsMovNro, gsOpeCod, txtMovDesc, _
                                        txtCuenta, sCtaHaber, pnImporte, psPersCod, _
                                        psTipoIF, psCtaIFCod, pnMovRef, txtFecha.Text, _
                                        .TextMatrix(.row, 15), .TextMatrix(.row, 16), .TextMatrix(.row, 17), IIf(.TextMatrix(.row, 15) = "3", 0, .TextMatrix(.row, 18)), _
                                        .TextMatrix(.row, 19), .TextMatrix(.row, 20), IIf(.TextMatrix(.row, 15) = "3", 0, .TextMatrix(.row, 21)), IIf(.TextMatrix(.row, 15) = "3", 0, .TextMatrix(.row, 22)), _
                                        .TextMatrix(.row, 23), .TextMatrix(.row, 24), .TextMatrix(.row, 25), .TextMatrix(.row, 4), lsMovNroACAdd) = 0 Then
                ImprimeAsientoContable lsMovNro
                If .TextMatrix(.row, 15) = 6 Then 'PASI20150921 ERS0712015
                    ImprimeAsientoContable lsMovNroACAdd
                End If
                If MsgBox("Desea Realizar otra Confirmación ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                    fgInversiones.EliminaFila fgInversiones.row
                    txtMovDesc = ""
                    Me.txtCuenta.Text = ""
                Else
                    Unload Me
                End If
            End If
        End With
    End If
End Sub
Private Sub cmdProcesar_Click()
 Dim oCaja As nCajaGeneral
 Dim rs As Recordset
 
 If ValidaFecha(Me.txtDesde.Text) <> "" Then
    MsgBox "Fecha de Inicio no Valida", vbInformation, "AVISO"
    Me.txtDesde.SetFocus
    Exit Sub
 ElseIf ValidaFecha(Me.txthasta.Text) <> "" Then
    MsgBox "Fecha Fin no Valida", vbInformation, "AVISO"
    Me.txthasta.SetFocus
    Exit Sub
 End If
 
 Set oCaja = New nCajaGeneral
 Set rs = New Recordset
 Set rs = oCaja.obtenerInversionesApertura(IIf(gsOpeCod = "421302", "421301", "422301"), Me.txtDesde.Text, Me.txthasta.Text, Trim(Right(Me.cmbTipo.Text, 2)))
 
fgInversiones.Clear
fgInversiones.FormaCabecera
fgInversiones.Rows = 2
Me.txtCuenta.Text = ""
Me.txtMovDesc.Text = ""
        
    If Not rs.EOF And Not rs.BOF Then
        Set fgInversiones.Recordset = rs
        fgInversiones.SetFocus
        cargarCuentasContables 'PASI20150921 ERS0712015
    Else
        txtCuenta.rs = New ADODB.Recordset 'PASI20150921 ERS0712015
        MsgBox "Datos no encontrados para proceso seleccionado", vbInformation, "Aviso"
    End If


RSClose rs
  
End Sub
Private Function Validar() As Boolean
    Validar = True
        If txtCuenta = "" Then
            MsgBox "Debe Seleccionar una Cuenta Contable", vbInformation, "¡Aviso!"
            Validar = False
            Me.txtCuenta.SetFocus
            Exit Function
        End If
        If Me.txtMovDesc.Text = "" Then
            MsgBox "Seleccione una Apertura de la Lista", vbInformation, "¡Aviso!"
            Validar = False
            Exit Function
        End If
End Function

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgInversiones_Click()
    With Me.fgInversiones
        If .TextMatrix(1, 1) <> "" Then
            Me.txtMovDesc.Text = .TextMatrix(.row, 12)
        End If
    End With
End Sub

Private Sub Form_Load()
    If gsOpeCod = "421302" Then
        Me.Caption = "Confirmacion de Inversiones MN"
    ElseIf gsOpeCod = "422302" Then
        Me.Caption = "Confirmacion de Inversiones ME"
    End If
    Me.txtFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
    txtDesde = DateAdd("d", (Day(gdFecSis) - 1) * -1, gdFecSis)
    txthasta = gdFecSis
    cargarTipoInversion
    'cargarCuentasContables 'Comentado xPASI, ERS0712015
End Sub
Private Sub cargarCuentasContables()
    Dim oCaja As New nCajaGeneral
    Dim rs As New Recordset
    'Set rs = oCaja.obtenerListaCtaCont(gsOpeCod, "1")
    Set rs = oCaja.getListaCtaContInversiones(Mid(gsOpeCod, 3, 1), IIf(Trim(Right(Me.cmbTipo.Text, 2)) = "%", 0, (Trim(Right(Me.cmbTipo.Text, 2))))) 'PASI20150921 ERS0712015 agrego el parametro 2
    
    
    If Not (rs.EOF And rs.BOF) Then
        txtCuenta.rs = rs
        txtCuenta.psRaiz = "Cuentas Contables"
    End If
    
    Set oCaja = Nothing
End Sub
Private Sub txtDesde_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        Me.txthasta.SetFocus
     End If
End Sub
Private Sub txthasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub
Private Sub cargarTipoInversion()
    Dim rsTpoInversion As ADODB.Recordset
    Dim oCons As DConstante
    
    Set rsTpoInversion = New ADODB.Recordset
    Set oCons = New DConstante
    
    Set rsTpoInversion = oCons.CargaConstante(9990)
    If Not (rsTpoInversion.EOF And rsTpoInversion.BOF) Then
        cmbTipo.Clear
        Do While Not rsTpoInversion.EOF
            cmbTipo.AddItem Trim(rsTpoInversion(2)) & Space(100) & Trim(rsTpoInversion(1))
            rsTpoInversion.MoveNext
        Loop
        cmbTipo.AddItem "Todos los Tipos" + Space(70) + "%", 0
        cmbTipo.ListIndex = 0
    End If
    
End Sub
