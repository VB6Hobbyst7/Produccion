VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmJoyGarRegistro 
   Caption         =   "Registro de Joyas para Garantía"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7725
   Icon            =   "frmJoyGarRegistro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7725
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraContenedor 
      Height          =   5055
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7455
      Begin VB.Frame fraContenedor 
         Height          =   1065
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   7215
         Begin VB.TextBox txtPiezas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   285
            Left            =   3480
            MaxLength       =   5
            TabIndex        =   15
            Top             =   600
            Width           =   1095
         End
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
            Height          =   270
            Left            =   1080
            TabIndex        =   16
            Top             =   540
            Width           =   1095
         End
         Begin VB.Label lblValorTasacion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5880
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "J. Bruto  (gr)"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "J. Neto  (gr)"
            Height          =   210
            Index           =   10
            Left            =   2520
            TabIndex        =   22
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Piezas"
            Height          =   210
            Index           =   2
            Left            =   2520
            TabIndex        =   21
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Tasación "
            Height          =   255
            Index           =   3
            Left            =   5040
            TabIndex        =   20
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblOroBruto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblOroNeto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3480
            TabIndex        =   18
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Cost. Custod."
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   990
         End
      End
      Begin VB.Frame fraPiezasDet 
         Caption         =   "Detalle de Piezas"
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   1620
         Width           =   7215
         Begin VB.CommandButton cmdPiezaEliminar 
            Caption         =   "-"
            Enabled         =   0   'False
            Height          =   495
            Left            =   6760
            TabIndex        =   7
            Top             =   1080
            Width           =   345
         End
         Begin VB.CommandButton CmdPiezaAgregar 
            Caption         =   "+"
            Height          =   495
            Left            =   6760
            TabIndex        =   6
            Top             =   480
            Width           =   345
         End
         Begin SICMACT.FlexEdit FEJoyas 
            Height          =   1695
            Left            =   135
            TabIndex        =   8
            Top             =   225
            Width           =   6555
            _ExtentX        =   11562
            _ExtentY        =   2990
            Cols0           =   8
            HighLight       =   1
            AllowUserResizing=   2
            EncabezadosNombres=   "Num-Pzas-Material-PBruto-PNeto-Tasac-Descripcion-Item"
            EncabezadosAnchos=   "400-450-1030-650-650-700-2500-0"
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-2-3-4-X-6-X"
            ListaControles  =   "0-0-3-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-L-R-R-R-L-C"
            FormatosEdit    =   "0-3-1-2-2-2-0-3"
            TextArray0      =   "Num"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Cliente"
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
         Height          =   1425
         Index           =   6
         Left            =   135
         TabIndex        =   34
         Top             =   120
         Width           =   6810
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            Height          =   345
            Left            =   180
            TabIndex        =   36
            Top             =   990
            Width           =   825
         End
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Enabled         =   0   'False
            Height          =   345
            Left            =   1080
            TabIndex        =   35
            Top             =   990
            Width           =   825
         End
         Begin MSComctlLib.ListView lstCliente 
            Height          =   795
            Left            =   90
            TabIndex        =   37
            Top             =   180
            Width           =   6555
            _ExtentX        =   11562
            _ExtentY        =   1402
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
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
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Codigo del Cliente"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre / Razón Social del Cliente"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Dirección"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Teléfono"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Ciudad"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Zona"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Doc.Civil"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Nro.Doc.Civil"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Doc.Tributario"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Nro.Doc.Tributario"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Kilataje"
         Height          =   1500
         Index           =   5
         Left            =   6000
         TabIndex        =   25
         Top             =   3720
         Visible         =   0   'False
         Width           =   1350
         Begin VB.TextBox txt21k 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   525
            MaxLength       =   5
            TabIndex        =   29
            Top             =   1140
            Width           =   720
         End
         Begin VB.TextBox txt18k 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   525
            MaxLength       =   5
            TabIndex        =   28
            Top             =   840
            Width           =   720
         End
         Begin VB.TextBox txt16k 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   540
            MaxLength       =   5
            TabIndex        =   27
            Top             =   540
            Width           =   720
         End
         Begin VB.TextBox txt14k 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   540
            MaxLength       =   5
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "16 K"
            Height          =   255
            Index           =   12
            Left            =   135
            TabIndex        =   33
            Top             =   585
            Width           =   420
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "21 K"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   32
            Top             =   1140
            Width           =   465
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "18 K"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "14 K"
            Height          =   210
            Index           =   11
            Left            =   120
            TabIndex        =   30
            Top             =   255
            Width           =   495
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   570
         Index           =   7
         Left            =   1920
         TabIndex        =   9
         Top             =   1680
         Width           =   5415
         Begin VB.Label lblEtiqueta 
            Caption         =   "Kilataje (gr)"
            Height          =   195
            Index           =   22
            Left            =   3120
            TabIndex        =   13
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Porcentaje (%)"
            Height          =   195
            Index           =   23
            Left            =   300
            TabIndex        =   12
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label lblOroPrestamo 
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
            Height          =   270
            Left            =   4080
            TabIndex        =   11
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label lblOroPrestamoPorcen 
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
            Height          =   255
            Left            =   1800
            TabIndex        =   10
            Top             =   240
            Width           =   1035
         End
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdImpVolTas 
      Caption         =   "&Volante de Tas."
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   5280
      Width           =   1335
   End
End
Attribute VB_Name = "frmJoyGarRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnTasaCustodia As Double
Dim fnPrecioOro14 As Double
Dim fnPrecioOro16 As Double
Dim fnPrecioOro18 As Double
Dim fnPrecioOro21 As Double
Dim fsContrato As String
Dim lnJoyas As Integer
Dim vOroBruto As Double
Dim vOroNeto As Double
Dim vPiezas As Integer
Dim vCostoCustodia As Double
'** lstCliente - ListView
Dim lstTmpCliente As ListItem
Dim lsIteSel As String

Private Sub CalculaCostosAsociados()
    Dim loCostos As COMNColoCPig.NCOMColPCalculos
    Set loCostos = New COMNColoCPig.NCOMColPCalculos
    vCostoCustodia = 0 'loCostos.nCalculaCostoCustodia(val(lblValorTasacion.Caption), fnTasaCustodia, val(cboPlazo.Text))
    Set loCostos = Nothing
    Me.lblCostoCustodia = Format(vCostoCustodia, "#0.00")
End Sub

Private Function ValorTasacion() As Double
If val(txt14k.Text) >= 0 And val(txt16k.Text) >= 0 And val(txt18k.Text) >= 0 And val(txt21k.Text) >= 0 Then
   ValorTasacion = (val(txt14k.Text) * fnPrecioOro14) + (val(txt16k.Text) * fnPrecioOro16) + (val(txt18k.Text) * fnPrecioOro18) + (val(txt21k.Text) * fnPrecioOro21)
Else
   MsgBox " No se ha ingresado correctamente el Kilataje ", vbInformation, " Aviso "
End If
End Function

Private Sub Limpiar()
    lblOroBruto.Caption = Format(0, "#0.00")
    lblOroNeto.Caption = Format(0, "#0.00")
    txtPiezas.Text = Format(0, "#0")
    lblValorTasacion.Caption = Format(0, "#0.00")
    Me.lblCostoCustodia = Format(0, "#0.00")
    txt14k.Text = Format(0, "#0.00")
    txt16k.Text = Format(0, "#0.00")
    txt18k.Text = Format(0, "#0.00")
    txt21k.Text = Format(0, "#0.00")
    Me.lblOroPrestamo.Caption = ""
    Me.lblOroPrestamoPorcen.Caption = ""
    lstCliente.ListItems.Clear
    FEJoyas.Clear
    FEJoyas.Rows = 2
    FEJoyas.FormaCabecera
    lnJoyas = 0
End Sub

Private Function SumaKilataje() As Double
If val(txt14k.Text) >= 0 And val(txt16k.Text) >= 0 And val(txt18k.Text) >= 0 And val(txt21k.Text) >= 0 Then
   SumaKilataje = val(txt14k.Text) + val(txt16k.Text) + val(txt18k.Text) + val(txt21k.Text)
Else
   MsgBox " No se ha ingresado correctamente el Kilataje ", vbInformation, " Aviso "
End If
End Function

Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
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

Exit Sub

ControlError:
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    txt14k.Enabled = False
    txt16k.Enabled = False
    txt18k.Enabled = False
    txt21k.Enabled = False
    txtPiezas.Enabled = False
    lblValorTasacion.Enabled = False
    cmdAgregar.Enabled = True
    cmdEliminar.Enabled = False
End Sub

Private Sub CmdGrabar_Click()
Dim pbTran As Boolean
Dim lrPersonas As New ADODB.Recordset
Dim lsMovNro As String
Dim lsFechaHoraGrab As String
Dim lnOroBruto As Double
Dim lnOroNeto As Double
Dim lnPiezas As Integer
Dim lnValTasacion As Currency
Dim lnCostoCustodia As Currency
Dim lrJoyas As New ADODB.Recordset
Dim loRegPig As COMNColoCPig.NCOMColPContrato
Dim loRegImp As COMNColoCPig.NCOMColPImpre
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim lsJoyGarCod As String
Dim loPrevio As previo.clsprevio
Dim lsCadImprimir As String
Dim lsmensaje As String

pbTran = False
If ValidaDatosGrabar = False Then Exit Sub
Set lrPersonas = fgGetCodigoPersonaListaRsNew(Me.lstCliente)
lnOroBruto = val(lblOroBruto.Caption)
lnOroNeto = val(lblOroNeto.Caption)
lnPiezas = val(txtPiezas.Text)
lnValTasacion = CCur(lblValorTasacion.Caption)
lnCostoCustodia = CCur(Me.lblCostoCustodia.Caption)

Set lrJoyas = FEJoyas.GetRsNew
If ValidarMsh = True Then Exit Sub
If MsgBox("¿Desea Registrar las Joyas ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
    Set loRegPig = New COMNColoCPig.NCOMColPContrato
        lsJoyGarCod = loRegPig.RegistrarJoyaGarantia(lrPersonas("cPersCod"), lnValTasacion, txtPiezas.Text, lblOroBruto.Caption, lblOroNeto.Caption, lsMovNro, lrJoyas)
        pbTran = False
    Set loRegPig = Nothing
    
    If MsgBox("Imprimir Registro de Joyas ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        Set loRegImp = New COMNColoCPig.NCOMColPImpre
            lsCadImprimir = ""
            lsCadImprimir = loRegImp.nPrintRegJoyaGarantia(lsJoyGarCod, lrPersonas, lsFechaHoraGrab, lnOroBruto, lnOroNeto, _
            lnValTasacion, lnPiezas, lsmensaje, gImpresora)
            If Trim(lsmensaje) <> "" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Sub
            End If
        Set loPrevio = New previo.clsprevio

        Dim oImp As New ContsImp.clsConstImp
            oImp.Inicia gImpresora
            gPrnSaltoLinea = oImp.gPrnSaltoLinea
            gPrnSaltoPagina = oImp.gPrnSaltoPagina
        Set oImp = Nothing
            loPrevio.PrintSpool sLpt, lsCadImprimir, False
            Do While True
            Dim cad As String
                If MsgBox("Reimprimir Registro de Joyas? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                            loPrevio.PrintSpool sLpt, gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & lsCadImprimir, False
                            loPrevio.PrintSpool sLpt, gPrnSaltoLinea
                Else
                    Exit Do
                End If
            Loop
    Set loRegImp = Nothing
    End If
End If

Set loPrevio = Nothing
Set loRegPig = Nothing
Limpiar
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
    Limpiar
End Sub

Private Sub cmdImpVolTas_Click()
Dim lrJoyas  As New ADODB.Recordset
Dim lnOroBruto As Double
Dim lnOroNeto As Double
Dim lnPiezas As Integer
Dim lnValTasacion As Currency
Dim lsCadImprimir As String
Dim loRegImp As COMNColoCPig.NCOMColPImpre
Dim loPrevio As previo.clsprevio

    If ValidarMsh = True Then Exit Sub
    Set lrJoyas = FEJoyas.GetRsNew
    lnOroBruto = val(lblOroBruto.Caption)
    lnOroNeto = val(lblOroNeto.Caption)
    lnPiezas = val(txtPiezas.Text)
    lnValTasacion = CCur(lblValorTasacion.Caption)
    If MsgBox("Imprimir Volante de Tasación ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        Set loRegImp = New COMNColoCPig.NCOMColPImpre
            lsCadImprimir = ""
            lsCadImprimir = loRegImp.ImprimeVolanteTasacionJoyas(lnOroBruto, lnOroNeto, lnValTasacion, lnPiezas, lrJoyas, gImpresora)
        Set loRegImp = Nothing
        Set loPrevio = New previo.clsprevio
        Dim oImp As New ContsImp.clsConstImp
            oImp.Inicia gImpresora
            gPrnSaltoLinea = oImp.gPrnSaltoLinea
            gPrnSaltoPagina = oImp.gPrnSaltoPagina
        Set oImp = Nothing
            loPrevio.PrintSpool sLpt, lsCadImprimir & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea, False
    End If
    Set loPrevio = Nothing
End Sub

Private Sub CmdPiezaAgregar_Click()
    If lstCliente.ListItems.Count = 0 Then
       MsgBox "No puede agregar items, si no figura(n) cliente(s) en el contrato.", vbOKOnly + vbInformation, "Atención"
       Exit Sub
    End If
    
    If FEJoyas.Rows <= 20 Then
        If lnJoyas > 1 And FEJoyas.TextMatrix(FEJoyas.Row, 6) = "" Then
            MsgBox "Ingrese datos de la Joya anterior", vbInformation, "Aviso"
            FEJoyas.SetFocus
            Exit Sub
        Else
            If lnJoyas = 1 And FEJoyas.TextMatrix(FEJoyas.Row, 6) = "" Then
                MsgBox "Ingrese datos de la Joya anterior", vbInformation, "Aviso"
                Exit Sub
            Else
                lnJoyas = lnJoyas + 1
                FEJoyas.AdicionaFila
                If FEJoyas.Rows >= 2 Then
                   cmdPiezaEliminar.Enabled = True
                End If
                
                Dim loConst As COMDConstantes.DCOMConstantes
                Dim lrMaterial As New ADODB.Recordset
                Set loConst = New COMDConstantes.DCOMConstantes
                
                FEJoyas.Col = 2
                Set lrMaterial = loConst.RecuperaConstantes(gColocPMaterialJoyas, , "C.cConsDescripcion")
                FEJoyas.CargaCombo lrMaterial
                Set lrMaterial = Nothing
                FEJoyas.Col = 1
                FEJoyas.SetFocus
            End If
        End If
    Else
        CmdPiezaAgregar.Enabled = False
        MsgBox "Sólo puede ingresar como máximo veinte piezas", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdPiezaEliminar_Click()
    FEJoyas.EliminaFila FEJoyas.Row
    If FEJoyas.Rows <= 20 Then
        CmdPiezaAgregar.Enabled = True
    End If
    lnJoyas = lnJoyas + 1
    SumaColumnas
End Sub

'Private Sub CmdPrevio_Click()
'    Dim pbTran As Boolean
'    Dim lsCtaReprestamo As String
'    Dim oImp As New COMNColoCPig.NCOMColPImpre
'
'    Dim lrPersonas As New ADODB.Recordset
'    Dim lsMovNro As String
'    Dim lsFechaHoraGrab As String
'    Dim lnMontoPrestamo As Currency
'    Dim lnPlazo As Integer
'    Dim lsFechaVenc As String
'    Dim lnOroBruto As Double
'    Dim lnOroNeto As Double
'    Dim lnPiezas As Integer
'    Dim lnValTasacion As Currency
'    Dim lsTipoContrato As String
'    Dim lsLote As String
'    Dim ln14k As Double, ln16k As Double, ln18k As Double, ln21k As Double
'    Dim lnIntAdelantado As Currency, lnCostoTasac As Currency, lnCostoCustodia As Currency, lnImpuesto As Currency
'
'    Dim loRegPig As COMNColoCPig.NCOMColPContrato
'
'    Dim loContFunct As COMNContabilidad.NCOMContFunciones
'
'    Dim lsContrato As String
'    Dim loPrevio As previo.clsprevio
'
'    Dim lsCadImprimir As String
'    On Error GoTo ControlError
'    pbTran = False
'
'    If ValidaDatosGrabar = False Then Exit Sub
'
'    Set lrPersonas = fgGetCodigoPersonaListaRsNew(Me.lstCliente)
'    lnMontoPrestamo = CCur(txtMontoPrestamo.Text)
'    lnPlazo = val(cboPlazo.Text)
'    lsFechaVenc = Format$(Me.lblFechaVencimiento, "mm/dd/yyyy")
'    lnValTasacion = CCur(lblValorTasacion.Caption)
'    lsTipoContrato = Switch(cboTipcta.ListIndex = 0, "I", cboTipcta.ListIndex = 1, "O", cboTipcta.ListIndex = 2, "Y")
'    lnIntAdelantado = CCur(Me.lblInteres.Caption)
'    lnCostoTasac = CCur(Me.lblCostoTasacion.Caption)
'    lnCostoCustodia = CCur(Me.lblCostoCustodia.Caption)
'    lnImpuesto = CCur(Me.lblImpuesto.Caption)
'
'    'Genera Mov Nro
'    Set loContFunct = New COMNContabilidad.NCOMContFunciones
'        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'    Set loContFunct = Nothing
'
'    lsFechaHoraGrab = Format(Now, "dd/mm/yyyy")     'hora
'
'    lsContrato = "INFORMATIVA"
'
'    lsCadImprimir = oImp.PrintHojaInformativa(lsContrato, lrPersonas, fnTasaInteresAdelantado, _
'        lnMontoPrestamo, lsFechaHoraGrab, Format(lsFechaVenc, "mm/dd/yyyy"), lnPlazo, lnOroBruto, lnOroNeto, lnValTasacion, _
'        lnPiezas, lsLote, ln14k, ln16k, ln18k, ln21k, lnIntAdelantado, lnCostoTasac, lnCostoCustodia, lnImpuesto, gsCodUser)
'    Set oImp = Nothing
'
'    Set loPrevio = New previo.clsprevio
'        loPrevio.PrintSpool sLpt, lsCadImprimir, False
'
'    Do While True
'        If MsgBox("Reimprimir Hoja Informativa ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
'            loPrevio.PrintSpool sLpt, lsCadImprimir, False
'        Else
'            Set loPrevio = Nothing
'            Set loRegPig = Nothing
'            Exit Do
'        End If
'    Loop
'
'    Set loPrevio = Nothing
'    Set loRegPig = Nothing
'
'    Exit Sub
'
'ControlError:   ' Rutina de control de errores.
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
'        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'    'Limpiar
'End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub FEJoyas_Click()
Dim loConst As COMDConstantes.DCOMConstantes
Dim lrMaterial As New ADODB.Recordset
Set loConst = New COMDConstantes.DCOMConstantes

Select Case FEJoyas.Col
Case 2
    Set lrMaterial = loConst.RecuperaConstantes(gColocPMaterialJoyas, , "C.cConsDescripcion")
    FEJoyas.CargaCombo lrMaterial
    Set lrMaterial = Nothing

End Select
Set loConst = Nothing
End Sub

Private Sub feJoyas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FEJoyas.Col = 1 Then
        If FEJoyas.TextMatrix(FEJoyas.Row, 7) <> "" Then
            CmdPiezaAgregar.SetFocus
            CmdPiezaAgregar_Click
        End If
    End If
End If
End Sub

Private Sub FEJoyas_OnCellChange(pnRow As Long, pnCol As Long)
Dim loColPCalculos As COMDColocPig.DCOMColPCalculos
Dim lnPOro As Double
    If FEJoyas.Col = 3 Then
        If FEJoyas.TextMatrix(FEJoyas.Row, 3) = "" Then
            MsgBox "Ingrese un Peso Bruto Correcto", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    If FEJoyas.Col = 4 Then     'Peso Neto

        If FEJoyas.TextMatrix(FEJoyas.Row, 4) <> "" Then
            If CCur(FEJoyas.TextMatrix(FEJoyas.Row, 4)) < 0 Then
                MsgBox "Peso Neto no puede ser negativo", vbInformation, "Aviso"
                FEJoyas.TextMatrix(FEJoyas.Row, 4) = 0
            Else
                If CCur(FEJoyas.TextMatrix(FEJoyas.Row, 4)) > CCur(FEJoyas.TextMatrix(FEJoyas.Row, 3)) Then
                    MsgBox "Peso Neto no puede ser mayor que Peso Bruto", vbInformation, "Aviso"
                    FEJoyas.TextMatrix(FEJoyas.Row, 4) = 0
                Else
                    'CalculaTasacion
                        Set loColPCalculos = New COMDColocPig.DCOMColPCalculos
                        lnPOro = loColPCalculos.dObtienePrecioMaterial(1, val(Right(FEJoyas.TextMatrix(FEJoyas.Row, 2), 3)), 1)
                        If lnPOro <= 0 Then
                            MsgBox "Precio del Material No ha sido ingresado en el Tarifario, actualice el Tarifario", vbInformation, "Aviso"
                            Exit Sub
                        End If
                        Set loColPCalculos = Nothing
                        'Calcula el Valor de Tasacion
                        FEJoyas.TextMatrix(FEJoyas.Row, 5) = Format$(val(FEJoyas.TextMatrix(FEJoyas.Row, 4) * lnPOro), "#####.00")
                            
                End If
            End If
        End If
        
    End If
    SumaColumnas
End Sub

Private Sub Form_Load()
    CargaParametros
    Limpiar
End Sub

'Valida el campo txtPiezas
Private Sub txtPiezas_GotFocus()
    fEnfoque txtPiezas
End Sub

Private Sub cmdEliminar_Click()
On Error GoTo ControlError
    Dim i As Integer, J As Integer
    If lstCliente.ListItems.Count = 0 Then
       MsgBox "No existen datos, imposible eliminar", vbInformation, "Aviso"
       cmdEliminar.Enabled = False
       lstCliente.SetFocus
       Exit Sub
    Else
       For i = 1 To lstCliente.ListItems.Count
           If lstCliente.ListItems.iTem(i) = lsIteSel Then
              lstCliente.ListItems.Remove (i)
              Exit For
           End If
       Next i
    End If
    lstCliente.SetFocus
    If lstCliente.ListItems.Count = 0 Then
        lblOroBruto = Format(0, "#0.00")
        cmdEliminar.Enabled = False
    ElseIf lstCliente.ListItems.Count = 1 Then
    End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub CmdAgregar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String
Dim liFil As Integer
Dim loColPFunc As COMDColocPig.DCOMColPFunciones

On Error GoTo ControlError
Set loPers = New COMDPersona.UCOMPersona
Set loPers = frmBuscaPersona.Inicio

If Not loPers Is Nothing Then
    lsPersCod = loPers.sPersCod

    '** Verifica que no este en lista
    For liFil = 1 To Me.lstCliente.ListItems.Count
        If lsPersCod = Me.lstCliente.ListItems.iTem(liFil).Text Then
           MsgBox " Cliente Duplicado ", vbInformation, "Aviso"
           Exit Sub
        End If
    Next liFil
    
    'Verifica si es Empleado
    If loPers.fgVerificaEmpleado(lsPersCod) = True Then
        MsgBox "El Cliente tambien es empleado de la CMAC,  " & _
               "NO puede tener Registro de Joyas", vbInformation, "Aviso"
        Exit Sub
    End If
    Set lstTmpCliente = lstCliente.ListItems.Add(, , lsPersCod)
        lstTmpCliente.SubItems(1) = loPers.sPersNombre
        lstTmpCliente.SubItems(2) = loPers.sPersDireccDomicilio
        lstTmpCliente.SubItems(3) = loPers.sPersTelefono
        lstTmpCliente.SubItems(6) = gPersIdDNI
        lstTmpCliente.SubItems(7) = loPers.sPersIdnroDNI
        lstTmpCliente.SubItems(9) = loPers.sPersIdnroRUC
    
        Set loColPFunc = New COMDColocPig.DCOMColPFunciones
            lstTmpCliente.SubItems(4) = Trim(loColPFunc.dObtieneNombreZonaPersona(loPers.sPersCod))
        Set loColPFunc = Nothing
    
    cmdEliminar.Enabled = True
End If
Set loPers = Nothing
If lstCliente.ListItems.Count >= 1 Then
    cmdAgregar.Enabled = False
End If
Exit Sub

ControlError:
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub lstCliente_ItemClick(ByVal iTem As MSComctlLib.ListItem)
    lsIteSel = iTem
End Sub

'Carga los Parametros
Private Sub CargaParametros()
    Dim loParam As COMDColocPig.DCOMColPCalculos
    Dim loConstSis As COMDConstSistema.NCOMConstSistema
    Set loParam = New COMDColocPig.DCOMColPCalculos
    fnTasaCustodia = loParam.dObtieneColocParametro(gConsColPTasaCustodia)
    Set loParam = Nothing
End Sub

Private Function ValidaDatosGrabar() As Boolean
Dim lbOk As Boolean
lbOk = True
If lstCliente.ListItems.Count <= 0 Then
    MsgBox "Falta ingresar el cliente" & vbCr & _
    " Cancele operación ", , " Aviso "
    lbOk = False
    Exit Function
End If
If val(lblOroBruto) < val(lblOroNeto) Then
    MsgBox " Oro Neto debe ser menor o igual a Oro Bruto ", vbInformation, " Aviso "
    lbOk = False
    Exit Function
End If
    
    If FEJoyas.Rows < 1 Then
        MsgBox " No ha ingresado el Detalle de las Joyas" & vbCr & " No se puede grabar con datos inconclusos ", vbInformation, " Aviso "
        cmdAgregar.SetFocus
        lbOk = False
        Exit Function
    Else
        FEJoyas.Row = 0
        txt14k.Text = 0: txt16k.Text = 0: txt18k.Text = 0: txt21k.Text = 0
        Do While FEJoyas.Row < FEJoyas.Rows - 1
            Select Case val(Right(FEJoyas.TextMatrix(FEJoyas.Row + 1, 2), 3))
                Case 14
                    txt14k.Text = val(txt14k.Text) + val(FEJoyas.TextMatrix(FEJoyas.Row + 1, 4))
                Case 16
                    txt16k.Text = val(txt16k.Text) + val(FEJoyas.TextMatrix(FEJoyas.Row + 1, 4))
                Case 18
                    txt18k.Text = val(txt18k.Text) + val(FEJoyas.TextMatrix(FEJoyas.Row + 1, 4))
                Case 21
                    txt21k.Text = val(txt21k.Text) + val(FEJoyas.TextMatrix(FEJoyas.Row + 1, 4))
            End Select
            FEJoyas.Row = FEJoyas.Row + 1
        Loop
    End If

ValidaDatosGrabar = lbOk
End Function

Private Sub SumaColumnas()
Dim i As Integer
Dim lnPiezasT As Integer, lnPBrutoT As Double, lnPNetoT As Double, lnTasacT As Double
    lnPiezasT = 0: lnPBrutoT = 0:       lnPNetoT = 0:       lnTasacT = 0 ':         lnPrestamoT = 0
    
    'TOTAL PIEZAS
    lnPiezasT = FEJoyas.SumaRow(1)
    txtPiezas.Text = Format$(lnPiezasT, "##")

    'PESO BRUTO
    lnPBrutoT = FEJoyas.SumaRow(3)
    lblOroBruto.Caption = Format$(lnPBrutoT, "######.00")

    'PESO NETO
    lnPNetoT = FEJoyas.SumaRow(4)
    lblOroNeto.Caption = Format$(lnPNetoT, "######.00")

    lnTasacT = FEJoyas.SumaRow(5)
    lblValorTasacion.Caption = Format$(lnTasacT, "######.00")
    cmdGrabar.Enabled = True
    cmdImpVolTas.Enabled = True
    CalculaCostosAsociados
End Sub

Private Function MostrarJoyasDet(ByVal prJoyas As ADODB.Recordset) As Boolean
    Dim i As Integer
    i = 1
    If prJoyas.BOF And prJoyas.EOF Then
        MsgBox " Error al mostrar datos del cliente ", vbCritical, " Aviso "
        MostrarJoyasDet = False
    Else
        Me.FEJoyas.rsFlex = prJoyas
        MostrarJoyasDet = True
    End If
End Function

Public Function ValidarMsh() As Boolean
    Dim nFilas As Integer
    Dim i As Integer
    nFilas = FEJoyas.Rows
    For i = 0 To nFilas - 1
        If FEJoyas.TextMatrix(i, 1) = "" Then
            ValidarMsh = True
            MsgBox "Ingrese el detalle de Joyas", vbInformation, "Aviso"
            Exit Function
        End If
    Next
End Function


