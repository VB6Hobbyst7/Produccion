VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmJoyGarRescate 
   Caption         =   "Rescate de Joya con Garantia"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   Icon            =   "frmJoyGarRescate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   8115
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar "
      Height          =   375
      Left            =   5760
      TabIndex        =   29
      Top             =   6945
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   6675
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtNroJoya 
         Height          =   285
         Left            =   5520
         TabIndex        =   36
         Top             =   300
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nro Joya"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Crédito"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   33
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   7080
         Picture         =   "frmJoyGarRescate.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Buscar ..."
         Top             =   240
         Width           =   420
      End
      Begin VB.Frame fraContenedor 
         Enabled         =   0   'False
         Height          =   1005
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   5520
         Width           =   7425
         Begin VB.Label lblCostoCustodia 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1800
            TabIndex        =   27
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label lblFecPago 
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1800
            TabIndex        =   26
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Costo Custodia"
            Height          =   225
            Index           =   18
            Left            =   360
            TabIndex        =   25
            Top             =   600
            Width           =   1380
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Dias Custodia"
            Height          =   225
            Index           =   19
            Left            =   3840
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblDiasTranscurridos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5160
            TabIndex        =   23
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fec.Cancelación"
            Height          =   225
            Index           =   16
            Left            =   345
            TabIndex        =   22
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   7455
         Begin VB.Label Label1 
            Caption         =   "Cliente:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Direccion:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Doc Ident."
            Height          =   255
            Left            =   4320
            TabIndex        =   18
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblCliente 
            Height          =   255
            Left            =   960
            TabIndex        =   17
            Top             =   240
            Width           =   5895
         End
         Begin VB.Label lblDireccion 
            Height          =   255
            Left            =   960
            TabIndex        =   16
            Top             =   600
            Width           =   3135
         End
         Begin VB.Label lblDocIdent 
            Height          =   255
            Left            =   5160
            TabIndex        =   15
            Top             =   600
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   7455
         Begin VB.Label Label4 
            Caption         =   "Piezas:"
            Height          =   255
            Left            =   600
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblPiezas 
            Height          =   255
            Left            =   1560
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Tasacion:"
            Height          =   255
            Left            =   4320
            TabIndex        =   11
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblTasacion 
            Height          =   255
            Left            =   5280
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "J. Bruto:"
            Height          =   255
            Left            =   4320
            TabIndex        =   9
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "J. Neto:"
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblJBruto 
            Height          =   255
            Left            =   5280
            TabIndex        =   7
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblJNeto 
            Height          =   255
            Left            =   1560
            TabIndex        =   6
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Descripcion Lote"
         Height          =   2055
         Left            =   120
         TabIndex        =   3
         Top             =   3360
         Width           =   7455
         Begin MSDataGridLib.DataGrid dgBuscar 
            Height          =   1575
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   2778
            _Version        =   393216
            AllowUpdate     =   0   'False
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   1
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "nItem"
               Caption         =   "Item"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "nPiezas"
               Caption         =   "Piezas"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "nPesoBruto"
               Caption         =   "J. Bruto"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "nPesoNeto"
               Caption         =   "J. Neto"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "cDescripcion"
               Caption         =   "Descripcion"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               Size            =   182
               BeginProperty Column00 
                  ColumnWidth     =   494.929
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   2940.095
               EndProperty
            EndProperty
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   3360
         TabIndex        =   28
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblNroJoya 
         Caption         =   "Nº Joya:"
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
         Left            =   3360
         TabIndex        =   35
         Top             =   300
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Buscar Por:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   6945
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   6945
      Width           =   975
   End
   Begin VB.Label lblCod 
      Caption         =   "Label5"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmJoyGarRescate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnMaxDiasCustodiaDiferida As Double
Dim fnTasaIGV As Double
Dim fnPorcentajeCustodiaDiferida As Double
Dim vCostoCustodiaExtra As Double
Dim vSaldoCustodiaExtra As Double
Dim fnVarOpeCod As Long

Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = gsCodCMAC & gsCodAge
    Set dgBuscar.DataSource = Nothing
    Me.lblCostoCustodia.Caption = Format(0, "#0.00")
    Me.lblFecPago.Caption = "  /  /    "
    Me.lblDiasTranscurridos.Caption = ""
    lblCliente.Caption = ""
    lblDireccion.Caption = ""
    lblDocIdent.Caption = ""
    lblPiezas.Caption = ""
    lblCod.Caption = ""
    lblTasacion.Caption = ""
    lblJNeto.Caption = ""
    lblJBruto.Caption = ""
    txtNroJoya.Text = ""
End Sub

Public Sub BuscaContrato(ByVal psNroContrato As String)
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim loCalculos As COMNColoCPig.NCOMColPCalculos
Dim lnCustodiaDiferida  As Currency
Dim lsmensaje As String
    
    Set lrValida = New ADODB.Recordset
        Set loValContrato = New COMNColoCPig.NCOMColPValida
        Set lrValida = loValContrato.nRescateJoyasGarant(psNroContrato, gsCodAge, gdFecSis, fnVarOpeCod, gsCodUser, Trim(txtNroJoya.Text), IIf((Option1(0).value = True), 1, 2), lsmensaje)
            If Trim(lsmensaje) <> "" Then
                 MsgBox lsmensaje, vbInformation, "Aviso"
                 Exit Sub
            End If
        
        lblPiezas.Caption = lrValida("nPiezas")
        lblCod.Caption = lrValida("cJoyGarCod")
        lblTasacion.Caption = Format(lrValida("nTasacion"), "#,#0.00")
        lblJNeto.Caption = lrValida("nPesoNeto")
        lblJBruto.Caption = lrValida("nPesoBruto")
            
        Set loValContrato = Nothing
    
        If lrValida Is Nothing Then
            Limpiar
            Set lrValida = Nothing
            Exit Sub
        End If
        
        Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
        Set loMuestraContrato = New COMDColocPig.DCOMColPContrato
        Dim lrCredPigPersonas As ADODB.Recordset
        Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCredJoyasPersonas(IIf((Option1(0).value = True), 1, 2), psNroContrato, Trim(txtNroJoya.Text))
        
        lblCliente.Caption = lrCredPigPersonas("cPersApellido") & " " & lrCredPigPersonas("cPersNombre")
        lblDireccion.Caption = Trim(lrCredPigPersonas("cPersDireccDomicilio"))
        lblDocIdent.Caption = IIf(IsNull(lrCredPigPersonas("NroDNI")), "", lrCredPigPersonas("NroDNI"))
        
        Dim lrJoyasDet As ADODB.Recordset
        Set lrJoyasDet = loMuestraContrato.dObtieneDatosCreditoGarantJoyasDet(IIf((Option1(0).value = True), 1, 2), psNroContrato, Trim(txtNroJoya.Text))
        
        If lrJoyasDet.RecordCount <> "0" Then
            dgBuscar.Visible = True
            Set dgBuscar.DataSource = lrJoyasDet
            dgBuscar.Refresh
            Screen.MousePointer = 0
        Else
            Set dgBuscar.DataSource = Nothing
            dgBuscar.Refresh
            dgBuscar.Visible = False
        End If
        
        If Option1(0).value = True Then
            Set loCalculos = New COMNColoCPig.NCOMColPCalculos
                lnCustodiaDiferida = loCalculos.nCalculaCostoCustodiaDiferida(lrValida!nTasacion, IIf(IsNull(lrValida!nDiasTranscurridos), 0, lrValida!nDiasTranscurridos), lrValida!nPorcentajeCustodia, lrValida!nTasaIGV)
            Set loCalculos = Nothing
            Me.lblCostoCustodia = Format(lnCustodiaDiferida - lrValida!nCustodiaDiferida, "#0.00")
            Me.lblFecPago = Format(lrValida!dCancelado, "dd/mm/yyyy")
            Me.lblDiasTranscurridos = IIf(IsNull(lrValida!nDiasTranscurridos), 0, lrValida!nDiasTranscurridos)
        End If
       
    Set lrValida = Nothing
        
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
        
    AXCodCta.Enabled = False
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Dim lrContratos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

    If Trim(lsPersCod) <> "" Then
        Set loPersContrato = New COMDColocPig.DCOMColPContrato
            Set lrContratos = loPersContrato.dObtieneListaCredJoyasPersona(lsPersCod, IIf((Option1(0).value = True), 1, 2))
        Set loPersContrato = Nothing
    End If
    
    Set loCuentas = New COMDPersona.UCOMProdPersona
        Set loCuentas = frmJoyGarPersona.Inicio(lsPersNombre, lrContratos, IIf((Option1(0).value = True), 1, 2))
        If loCuentas.sCtaCod <> "" Then
            If Option1(0).value = True Then
                AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
                AXCodCta.SetFocusCuenta
            Else
                txtNroJoya.Text = Trim(loCuentas.sCtaCod)
                txtNroJoya.SetFocus
            End If
        End If
    Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    cmdGrabar.Enabled = False
    AXCodCta.Enabled = True
End Sub

Private Sub cmdGrabar_Click()
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarResc As COMNColoCPig.NCOMColPContrato
Dim oImp As New COMNColoCPig.NCOMColPImpre
Dim loColPRes As COMDColocPig.DCOMColPContrato
Dim loPrevio As previo.clsprevio

Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lsCuenta As String
Dim lscadimp As String

Dim lnSaldoCap As Currency, lnInteresComp As Currency, lnImpuesto As Currency
Dim lnCostoTasacion As Currency, lnCostoCustodia As Currency
Dim lnMontoEntregar As Currency

If MsgBox(" Grabar Rescate de Joya? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarResc = New COMNColoCPig.NCOMColPContrato

            Call loGrabarResc.nRescataJoyaGarantia(Trim(lblCod.Caption), 3, lsMovNro)
                                  
        Set loGrabarResc = Nothing

        'Impresión
        If MsgBox("Desea realizar impresiones ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            Set loPrevio = New previo.clsprevio
            lscadimp = oImp.ImprimirRescateJoyGar(AXCodCta.NroCuenta, lblCliente.Caption, 0, 0, gdFecSis, gsCodUser, gImpresora)
            loPrevio.PrintSpool sLpt, lscadimp, False, 22
            Do While True
                If MsgBox("Desea reimprimir ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                     loPrevio.PrintSpool sLpt, lscadimp, False, 22
                Else
                    Exit Do
                    Set loPrevio = Nothing
                End If
            Loop
            Set loPrevio = Nothing
        End If
        Limpiar
        Set oImp = Nothing
        AXCodCta.Enabled = True
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub CargaParametros()
    Dim loParam As COMDColocPig.DCOMColPCalculos
    Set loParam = New COMDColocPig.DCOMColPCalculos
    fnMaxDiasCustodiaDiferida = loParam.dObtieneColocParametro(gConsColPMaxDiasCustodiaDiferida)
    fnTasaIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)
    fnPorcentajeCustodiaDiferida = loParam.dObtieneColocParametro(gConsColPPorcentajeCustodiaDiferida)
    Set loParam = Nothing
End Sub

Private Sub Form_Load()
    fnVarOpeCod = gCredOpeDevJoyaGarantia
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Limpiar
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).value = True Then
    AXCodCta.Visible = True
    lblNroJoya.Visible = False
    txtNroJoya.Visible = False
    cmdGrabar.Enabled = False
    AXCodCta.Enabled = True
End If
If Option1(1).value = True Then
    cmdGrabar.Enabled = False
    AXCodCta.Visible = False
    lblNroJoya.Visible = True
    txtNroJoya.Visible = True
End If
Limpiar
End Sub

Private Sub txtNroJoya_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
End Sub
