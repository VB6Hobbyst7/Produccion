VERSION 5.00
Begin VB.Form frmCapCambioTasa 
   Caption         =   "Cambio de Tasa y/o Tasa Pactada"
   ClientHeight    =   4110
   ClientLeft      =   4590
   ClientTop       =   3330
   ClientWidth     =   8025
   Icon            =   "frmCapCambioTasa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8025
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   7575
         Begin VB.CheckBox chkTasaPactada 
            Caption         =   "Tasa Pactada"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3240
            TabIndex        =   17
            Top             =   2760
            Width           =   1815
         End
         Begin SICMACT.EditMoney txtTasa 
            Height          =   330
            Left            =   6720
            TabIndex        =   14
            Top             =   2280
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.CommandButton cmdCancelar 
            Height          =   375
            Left            =   4200
            Picture         =   "frmCapCambioTasa.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   495
         End
         Begin SICMACT.FlexEdit grdCliente 
            Height          =   1455
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   2566
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Código-Nombre-Relación-a-b"
            EncabezadosAnchos=   "300-1700-3500-1500-0-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColor       =   16777215
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "R-L-L-L-L-L"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   300
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
            CellBackColor   =   16777215
         End
         Begin VB.CommandButton cmdBuscarCta 
            Caption         =   "..."
            Height          =   375
            Left            =   3720
            TabIndex        =   3
            Top             =   240
            Width           =   375
         End
         Begin SICMACT.ActXCodCta txtCuenta 
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   661
            Texto           =   "Cuenta N°"
            EnabledCMAC     =   -1  'True
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
            EnabledAge      =   -1  'True
         End
         Begin VB.Label lblMoneda 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   960
            TabIndex        =   16
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "Moneda"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label lblTasa 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   330
            Left            =   4200
            TabIndex        =   13
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label lblTipoTasa 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   330
            Left            =   960
            TabIndex        =   12
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo tasa"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Nueva tasa"
            Height          =   255
            Left            =   5760
            TabIndex        =   10
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Tasa actual"
            Height          =   375
            Left            =   3240
            TabIndex        =   9
            Top             =   2280
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmCapCambioTasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nProducto As Producto
Dim bEdit As Boolean
Dim lbTasaPactada As Boolean
Dim lnTasaIntNominal As Double
Dim lbCambioTasa As Boolean
'By capi 21012009
Dim objPista As COMManejador.Pista



'By Capi 07082008
'Public Sub Inicia(ByVal nProd As Producto)
Public Sub Inicia(ByVal nProd As Producto, Optional ByVal bCambioTasa As Boolean = True)
    nProducto = nProd
     Select Case nProducto
     Case gCapAhorros
        txtCuenta.Prod = Trim(str(gCapAhorros))
        Me.Caption = Me.Caption & " Ahorro Corriente"
     Case gCapPlazoFijo
        txtCuenta.Prod = Trim(str(gCapPlazoFijo))
        Me.Caption = Me.Caption & " Plazo Fijo"
        'By capi 21012009
            If bCambioTasa Then
                gsOpeCod = gPFCambioTasaInteres
            Else
                gsOpeCod = gPFPactoTasaInteres
            End If
            
        '
     Case gCapCTS
        txtCuenta.Prod = Trim(str(gCapCTS))
        Me.Caption = Me.Caption & " CTS"
     End Select
     lbCambioTasa = bCambioTasa
     txtCuenta.CMAC = gsCodCMAC
     txtCuenta.EnabledCMAC = False
     txtCuenta.EnabledProd = False
     cmdEditar.Enabled = False
     cmdGrabar.Enabled = False
     cmdCancelar.Enabled = False
     lblTasa.Caption = ""
     txtTasa.Text = ""
     txtTasa.Enabled = False
     chkTasaPactada.Enabled = False
     bEdit = False
     Me.Show vbModal
End Sub

Private Sub cmdBuscarCta_Click()
    Dim clsPers As COMDPersona.UCOMPersona

    Set clsPers = New COMDPersona.UCOMPersona
    Set clsPers = frmBuscaPersona.Inicio

    If Not clsPers Is Nothing Then
        Dim sPers As String
        Dim rsPers As New ADODB.Recordset
        Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
        Dim sCta As String
        Dim sRelac As String * 15
        Dim sEstado As String
        Dim clsCuenta As UCapCuenta
        sPers = clsPers.sPersCod
        Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsPers = clsCap.GetCuentasPersona(sPers, nProducto)
        Set clsCap = Nothing
        If Not (rsPers.EOF And rsPers.EOF) Then
            Do While Not rsPers.EOF
                sCta = rsPers("cCtaCod")
                sRelac = rsPers("cRelacion")
                sEstado = Trim(rsPers("cEstado"))
                frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
                rsPers.MoveNext
            Loop
            Set clsCuenta = New UCapCuenta
            Set clsCuenta = frmCapMantenimientoCtas.Inicia
            If clsCuenta Is Nothing Then
            Else
                If clsCuenta.sCtaCod <> "" Then
                    txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
                    txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
                    txtCuenta.SetFocusCuenta
                    SendKeys "{Enter}"
                End If
            End If
            Set clsCuenta = Nothing
        Else
            Dim Mensaje As String
            Mensaje = "Persona no posee ninguna cuenta de captaciones "
            Select Case nProducto
                Case gCapAhorros
                    Mensaje = Mensaje & "Ahorro Corriente."
                Case gCapPlazoFijo
                    Mensaje = Mensaje & "Plazo Fijo."
                Case gCapCTS
                    Mensaje = Mensaje & "CTS."
            End Select
            MsgBox Mensaje, vbInformation, "Aviso"
        End If
        rsPers.Close
        Set rsPers = Nothing
    End If
    Set clsPers = Nothing
    txtCuenta.SetFocusCuenta
End Sub

Private Sub cmdCancelar_Click()
    cmdEditar.Caption = "Editar"
    cmdEditar.Enabled = False
    cmdGrabar.Enabled = False
    txtTasa.Enabled = False
    cmdBuscarCta.Enabled = True
    grdCliente.Clear
    txtCuenta.Age = ""
    txtCuenta.Cuenta = ""
    txtTasa.Text = ""
    txtTasa.Enabled = False
    lblTipoTasa.Caption = ""
    lblTasa.Caption = ""
    lblMoneda.Caption = ""
    lblMoneda.BackColor = &HFFFFFF
    txtCuenta.SetFocusAge
    bEdit = False
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub CmdEditar_Click()
    If bEdit = False Then
        cmdEditar.Caption = "Cancelar"
        cmdGrabar.Enabled = True
        'By Capi 07082008
        If lbCambioTasa = True Then
            chkTasaPactada.value = 1
            txtTasa.Enabled = True
            txtTasa.SetFocus
        Else
            txtTasa.value = lblTasa.Caption
            txtTasa.Enabled = False
            chkTasaPactada.Enabled = True
            If lbTasaPactada = True Then
                chkTasaPactada.value = 1
            Else
                chkTasaPactada.value = 0
            End If
        End If
        '
        bEdit = True
    Else
        cmdEditar.Caption = "Editar"
        cmdGrabar.Enabled = False
        txtTasa.Text = ""
        txtTasa.Enabled = False
        chkTasaPactada.Enabled = False
        bEdit = False
    End If
End Sub

Private Sub CmdGrabar_Click()
    If CDbl(txtTasa.Text) <= 0 Then
        MsgBox "La tasa debe ser mayor que cero.", vbCritical, "SICMACM"
        Exit Sub
    End If
    If MsgBox("Esta seguro de Registrar", vbQuestion + vbYesNo, "SICMACM") = vbNo Then
        Exit Sub
    End If
    
    'On Error GoTo CambioTasa
    
    Dim oCap As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim oMov As COMDMov.DCOMMov
    Dim VCMovNro As String
    Dim sCuenta As String
    Dim nTasa As Double
    Set oCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set oMov = New COMDMov.DCOMMov
    VCMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    'By Capi 13082008 para calculos segun los valores de tasas
    If lblTasa.Caption <> txtTasa.Text Then
        nTasa = ConvierteTEAaTNA(CDbl(txtTasa.Text))
    Else
        nTasa = lnTasaIntNominal
    End If
    sCuenta = txtCuenta.NroCuenta
    If chkTasaPactada = 1 Then
        lbTasaPactada = True
    Else
        lbTasaPactada = False
    End If
    oCap.CambioTasa sCuenta, CDbl(lblTasa.Caption), nTasa, VCMovNro, lbTasaPactada, lbCambioTasa
    'By Capi 21012009
    If lbCambioTasa Then
     objPista.InsertarPista gsOpeCod, VCMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Cambio Tasa"
    Else
     objPista.InsertarPista gsOpeCod, VCMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Tasa Pactada"
    End If
    '
            
    cmdCancelar_Click
    
'CambioTasa:
 '   Err.Raise Err.Number, "Error", Err.Description
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCta As String
        sCta = txtCuenta.NroCuenta
        ObtieneDatosCuenta sCta
    End If
    'By Capi 20012009
    Set objPista = New COMManejador.Pista
    '
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsCta As ADODB.Recordset, rsRel As ADODB.Recordset
    Dim nEstado As COMDConstantes.CaptacEstado
    Dim ssql As String

    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    
    If Not (rsCta.EOF And rsCta.BOF) Then
        nEstado = rsCta("nPrdEstado")
        lblMoneda = IIf(Mid(sCuenta, 9, 1) = "1", "Nacional", "Extranjera")
        lblMoneda.BackColor = IIf(Mid(sCuenta, 9, 1) = "1", &HFFFFFF, &HC0FFC0)
        'By Capi 07082008
        'lblTasa = Format$(ConvierteTNAaTEA(rsCta("nTasaInteres")), "#0.000")
        lnTasaIntNominal = rsCta("nTasaInteres")
        lblTasa = Round(Format$(ConvierteTNAaTEA(rsCta("nTasaInteres")), "#0.000"), 2)
        lblTipoTasa = rsCta("cTipoTasa")
        'By Capi 07082008
        lbTasaPactada = rsCta("bTasaPactada")
        rsCta.Close
        Set rsCta = clsMant.GetProductoPersona(sCuenta)
        If Not (rsCta.EOF And rsCta.BOF) Then
            Set grdCliente.Recordset = rsCta
            cmdBuscarCta.Enabled = False
            cmdCancelar.Enabled = True
            If (nEstado <> gCapEstActiva) Then
                MsgBox "Esta cuenta no se podrá editar por tener un estado no activa.", vbInformation, "Aviso"
                cmdEditar.Enabled = False
            Else
                cmdEditar.Enabled = True
                cmdEditar.SetFocus
            End If
        Else
            MsgBox "Cuenta no posee relacion con Persona", vbExclamation, "Aviso"
            txtCuenta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub
