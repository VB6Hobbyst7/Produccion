VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenNegocioBancos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   11070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin Sicmact.FlexEdit flxPenBanc 
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5741
      Cols0           =   15
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmCajaGenNegocioBancos.frx":0000
      EncabezadosAnchos=   "0-600-1800-3000-2200-1600-2400-3200-0-0-0-0-0-0-0"
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
      ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-4-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-L-R-C-L-C-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-2-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CheckBox chkMonto 
      Caption         =   "Monto"
      Height          =   195
      Left            =   8040
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   7920
      TabIndex        =   34
      Top             =   120
      Width           =   1470
      Begin VB.TextBox txtMonto 
         Enabled         =   0   'False
         Height          =   330
         Left            =   200
         TabIndex        =   1
         Top             =   220
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdRegulariza 
      Caption         =   "Regulariza"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   5760
      Width           =   1275
   End
   Begin VB.CheckBox chkRegulariza 
      Caption         =   "Regularizar"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   3360
      TabIndex        =   33
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Excel"
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   5880
      Width           =   1275
   End
   Begin VB.ComboBox cboComision 
      Height          =   315
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4950
      Visible         =   0   'False
      Width           =   3270
   End
   Begin VB.Frame FraAreas 
      Height          =   645
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   5400
      Begin VB.CheckBox chktodas 
         Caption         =   "Todas las Agencias"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   1710
      End
      Begin Sicmact.TxtBuscar txtBuscarArea 
         Height          =   330
         Left            =   1080
         TabIndex        =   3
         Top             =   210
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   582
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
         ForeColor       =   8388608
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agencias :"
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
         Height          =   195
         Left            =   135
         TabIndex        =   31
         Top             =   270
         Width           =   915
      End
      Begin VB.Label lblAreaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2235
         TabIndex        =   30
         Top             =   210
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8400
      TabIndex        =   14
      Top             =   5880
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9660
      TabIndex        =   15
      Top             =   5880
      Width           =   1275
   End
   Begin VB.Frame FraTipoCambio 
      Caption         =   "Tipo Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   630
      Left            =   120
      TabIndex        =   24
      Top             =   5640
      Width           =   3015
      Begin VB.TextBox txtTCBanco 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1905
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fijo:"
         Height          =   195
         Left            =   105
         TabIndex        =   28
         Top             =   270
         Width           =   285
      End
      Begin VB.Label lblTCFijo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   465
         TabIndex        =   27
         Top             =   225
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   1365
         TabIndex        =   26
         Top             =   255
         Width           =   510
      End
   End
   Begin VB.Frame frmBanco 
      Caption         =   "Cuenta Institucion Financiera"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   6255
      Begin Sicmact.TxtBuscar txtBuscarBanco 
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   582
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
         sTitulo         =   ""
      End
      Begin VB.Label lblDescCtabanco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   23
         Top             =   660
         Width           =   6015
      End
      Begin VB.Label lblDescbanco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2205
         TabIndex        =   22
         Top             =   285
         Width           =   3930
      End
   End
   Begin VB.Frame fradoc 
      Caption         =   "&Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   6480
      TabIndex        =   17
      Top             =   4320
      Width           =   4455
      Begin VB.TextBox txtNroDoc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   10
         Top             =   650
         Width           =   1695
      End
      Begin VB.ComboBox cboDocumento 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   250
         Width           =   4215
      End
      Begin MSMask.MaskEdBox txtFechaDoc 
         Height          =   330
         Left            =   3100
         TabIndex        =   18
         Top             =   640
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblComi 
         Caption         =   "Comisión"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   690
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lblDocFec 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   2500
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblDocNro 
         AutoSize        =   -1  'True
         Caption         =   "Nº :"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   270
      End
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   330
      Left            =   6450
      TabIndex        =   4
      Top             =   240
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
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
      Left            =   5760
      TabIndex        =   16
      Top             =   285
      Width           =   600
   End
End
Attribute VB_Name = "frmCajaGenNegocioBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oDocRec As NDocRec
Dim oOpe As DOperacion
Dim lbDeposito As Boolean
Dim oCtaIf As NCajaCtaIF
Dim lsCtaContBanco As String
Dim lsAjusteCta As String
Dim lnMontoDif As Currency
Dim oCaja  As nCajaGeneral
Dim lsCtaPendiente As String

Dim fnMovNroCap As Long
Dim fcOpeCodCap, fcNroDocCap, fcRefenciaCap, fcDocDesCap, fcDocFecha As String '***Agregado por ELRO el 20121001, según OYP-RFC106-2012
Dim objPista As COMManejador.Pista 'ARLO20170217

Dim bProcesado As Boolean 'ande 21070902 ERS029-2017
Public Sub inicio(Optional pbDeposito As Boolean = True)
lbDeposito = pbDeposito
Me.Show 1
End Sub

Public Sub LlenarGrid()
    'esperando
    cmdProcesar.Caption = "Espere..."
    cmdProcesar.Enabled = False

    'limpiar flex
    Call LimpiarFlexPenBanc
    
    If chktodas.value = 1 Then
        CargaPendientesNegocio ""
    Else
        CargaPendientesNegocio Mid(txtBuscarArea, 4, 2)
    End If
    
    'ya no espera
    cmdProcesar.Caption = "Procesar"
    cmdProcesar.Enabled = True
    bProcesado = True
End Sub

Private Sub cboDocumento_Click()
MuestraControlDoc True
If cboDocumento <> "" Then
    Select Case Val(Right(cboDocumento, 2))
        Case TpoDocCarta, TpoDocCheque, TpoDocOrdenPago, TpoDocNotaAbono, TpoDocNotaCargo
            MuestraControlDoc False
    End Select
End If
If Left(cboDocumento, 7) = "NINGUNO" Then
    txtNroDoc.Enabled = False
End If
End Sub

Private Sub MuestraControlDoc(pbActiva As Boolean)
txtNroDoc.Enabled = pbActiva
txtNroDoc.Visible = pbActiva
txtFechaDoc.Visible = pbActiva
lblDocFec.Visible = pbActiva
lblDocNro.Visible = pbActiva
lblComi.Visible = Not pbActiva
cboComision.Visible = Not pbActiva
End Sub

Private Sub Check1_Click()
txtMonto.Enabled = IIf(chkMonto.value = 1, True, False)
txtMonto.Text = ""
End Sub

Private Sub chkMonto_Click()
txtMonto.Enabled = IIf(chkMonto.value = 1, True, False)
txtMonto.Text = ""
End Sub

Private Sub chkRegulariza_Click()
If chkRegulariza.value = 1 Then
    cmdRegulariza.Enabled = True
Else
    cmdRegulariza.Enabled = False
End If
End Sub

Private Sub chktodas_Click()
If chktodas.value = 1 Then
    txtBuscarArea = ""
    lblAreaDesc = ""
    txtBuscarArea.Enabled = False
    
Else
    'lswPendBanc.ListItems.Clear
    txtBuscarArea.Enabled = True
End If

End Sub

Private Sub cmdAceptar_Click()
On Error GoTo AceptarErr
'Modificado por ande 20170901 ERS029-2017
'obtener los item selecccionados
Dim nRows, i, nTotalChecks As Integer
nRows = flxPenBanc.Rows
nTotalChecks = 0
For i = 0 To nRows - 1
    If flxPenBanc.TextMatrix(i, 1) = "." Then
        nTotalChecks = nTotalChecks + 1
    End If
Next i
If bProcesado = False Then
    MsgBox "Primero debe procesar los datos.", vbOKOnly + vbInformation, "Aviso"
    Exit Sub
End If
If nTotalChecks = 0 Then
    MsgBox "Debe seleccionar por lo menos un registro.", vbOKOnly + vbInformation, "Aviso"
    Exit Sub
End If
'declarando array
Dim aRegistros() As Variant
Dim iRow As Integer
iRow = 0
ReDim aRegistros(nTotalChecks, 13)
'llenandos array con datos de los registros seleccionados
For i = 0 To nRows - 1
    If flxPenBanc.TextMatrix(i, 1) = "." Then
        aRegistros(iRow, 0) = flxPenBanc.TextMatrix(i, 2) 'fecha
        aRegistros(iRow, 1) = flxPenBanc.TextMatrix(i, 3) 'operación
        aRegistros(iRow, 2) = flxPenBanc.TextMatrix(i, 4) 'banco
        aRegistros(iRow, 3) = flxPenBanc.TextMatrix(i, 5) 'monto
        aRegistros(iRow, 4) = flxPenBanc.TextMatrix(i, 6) 'agencia
        aRegistros(iRow, 5) = flxPenBanc.TextMatrix(i, 7) 'glosa
        aRegistros(iRow, 6) = flxPenBanc.TextMatrix(i, 8) 'cuenta contable
        aRegistros(iRow, 7) = flxPenBanc.TextMatrix(i, 9) 'Cod Agencia
        aRegistros(iRow, 8) = flxPenBanc.TextMatrix(i, 10) 'Nro Mov
        aRegistros(iRow, 9) = flxPenBanc.TextMatrix(i, 11) 'Cod Operación
        aRegistros(iRow, 10) = flxPenBanc.TextMatrix(i, 12) 'Referencia
        aRegistros(iRow, 11) = flxPenBanc.TextMatrix(i, 13) 'Nro Doc
        aRegistros(iRow, 12) = flxPenBanc.TextMatrix(i, 14) 'Fecha Doc
        iRow = iRow + 1
    End If
Next i

'obteniendo el total de registros seleccionados
nRows = UBound(aRegistros) - 1
'Variable de mensaje

Dim cMensaje As String
If nRows > 1 Then
    cMensaje = "¿Desea grabar las operaciones seleccionadas?"
Else
    cMensaje = "¿Desea grabar la operación seleccionada?"
End If

'variables para impresion del asiento contable
Dim cImpresion As String, cMovNro_largo As String, cTpoDoc As String, cDocumento As String, cDocNotaAC As String
cMovNro_largo = ""
cTpoDoc = ""
cDocumento = ""
cDocNotaAC = ""
'end ande

'***Modificado por ELRO el 20121002, según OYP-RFC106-2012
If fcOpeCodCap = "200103" Or fcOpeCodCap = "200203" Or fcOpeCodCap = "210103" Or _
   fcOpeCodCap = "210807" Or fcOpeCodCap = "220103" Or fcOpeCodCap = "220203" Then
    Dim oNContFunciones As New NContFunciones
    Dim lsMovNro2, lsGlosa2 As String
    Dim lnConfirmar As Long
    Dim lnImporte2 As Currency
    
    If Valida = False Then Exit Sub
       
    If MsgBox(cMensaje, vbYesNo + vbQuestion, "Aviso") = vbYes Then
        
        cmdAceptar.Caption = "Guardando..." 'ande 20170908 ERS029-2017
        cmdAceptar.Enabled = False
        cmdCancelar.Enabled = False
        
        For i = 0 To nRows
            'ANDE 20170829 ERS029-2017
            lsMovNro2 = oNContFunciones.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
            'lsGlosa2 = lswPendBanc.ListItems.Item(lswPendBanc.SelectedItem.Index).SubItems(5)
            lsGlosa2 = aRegistros(i, 5)
            'lnImporte2 = CCur(lswPendBanc.ListItems.Item(lswPendBanc.SelectedItem.Index).SubItems(3))
            lnImporte2 = aRegistros(i, 3)
            'END ANDE
                    
            lnConfirmar = oCaja.registrarOperacionesTranferenciaNegocio(lsMovNro2, gsOpeCod, lsGlosa2, fnMovNroCap, lnImporte2, nVal(lblTCFijo))
            
            If lnConfirmar > 0 Then
                'ImprimeAsientoContable lsMovNro2, , "55", fcNroDocCap
                cMovNro_largo = cMovNro_largo & lsMovNro2 & ","
            Else
                MsgBox "No se registró la operación.", vbInformation, "!Aviso¡"
                cmdAceptar.Caption = "Aceptar"
                cmdAceptar.Enabled = True
                cmdCancelar.Enabled = True
                'Exit Sub
                GoTo Salir
            End If
            
        Next i
        
        cmdAceptar.Enabled = True
        cmdCancelar.Enabled = True
            
            If MsgBox("Desea ingresar otra operación?", vbQuestion + vbYesNo, "¡Confirmacion!") = vbNo Then
                lsMovNro2 = ""
                lsGlosa2 = ""
                txtBuscarBanco = ""
                lblDescbanco = ""
                lblDescCtabanco = ""
                cboDocumento.ListIndex = -1
                txtNroDoc = ""
                txtFechaDoc = gdFecSis
                frmBanco.Enabled = True
                fradoc.Enabled = True
                Frame1.Enabled = True
                chkRegulariza.Enabled = True
                fnMovNroCap = 0
                fcOpeCodCap = ""
                fcRefenciaCap = ""
                fcNroDocCap = ""
                fcDocFecha = ""
                Unload Me
            Else
                lsMovNro2 = ""
                lsGlosa2 = ""
                txtBuscarBanco = ""
                lblDescbanco = ""
                lblDescCtabanco = ""
                cboDocumento.ListIndex = -1
                txtNroDoc = ""
                txtFechaDoc = gdFecSis
                frmBanco.Enabled = True
                fradoc.Enabled = True
                Frame1.Enabled = True
                chkRegulariza.Enabled = True
                fnMovNroCap = 0
                fcOpeCodCap = ""
                fcRefenciaCap = ""
                fcNroDocCap = ""
                fcDocFecha = ""
                'lswPendBanc.SelectedItem.Delete
            End If
        
        cMovNro_largo = Left(cMovNro_largo, Len(cMovNro_largo) - 1)
        If nRows > 1 Then
            ImprimeAsientoContable cMovNro_largo, "55", , , , , , , , , , , , , , , , , , , , True
        Else
            ImprimeAsientoContable cMovNro_largo, "55", , , , , , , , , , , , , , , , , , , , False
        End If
        
    End If
Else
    Dim lsNroDoc As String
    Dim lnTpoDoc As TpoDoc
    Dim lsDocumento As String
    Dim lsDocNotaAC As String
    Dim lsNroVoucher As String
    Dim lsNroNotaAC As String
    Dim lnMotivoNAC As MotivoNotaAbonoCargo
    Dim lsObjetoPadre As String
    Dim lsObjetoCod As String
    
    Dim lsCadBol As String
    Dim lnMotivoNACAux As MotivoNotaAbonoCargo
    Dim lsObjetoPadreAux As String
    Dim lsObjetoCodAux As String
    
    Dim oDoc As clsDocPago
    Dim oCont As NContFunciones
    Dim oDocRec As NDocRec
    Dim oContImp As NContImprimir
    Dim lsCuentaAho As String
    Dim lnTpoDocAux As TpoDoc
    
    Dim lsMovNro As String
    Dim lsPersNombre As String
    Dim lsPersDireccion As String
    Dim lsUbigeo As String
    Dim lsCuentaAhoAux As String
    Dim lnMontoAux As Currency
    Dim lsEntiOrig As String
    Dim lsCtaEntOrig As String
    Dim lsEntiDest As String
    Dim lsCtaEntDest As String
    Dim lsSubCtaIF   As String
    
    Set oDocRec = New NDocRec
    Set oContImp = New NContImprimir
    Set oDoc = New clsDocPago
    Dim lsGlosa As String
    Set oOpe = New DOperacion
    
    Dim lbGrabaNegocio As Boolean
    
    Dim lnImporte As Currency
    Dim lnMovnroRef As Long
    Dim lnCtaPlantillaNeg As String
    Dim R2 As ADODB.Recordset
    Set R2 = New ADODB.Recordset
    Dim lsAgeCod As String
    Dim lnCtaCompara As String
    
    '***Agregado por ELRO el 20120112, según Acta N° 003-2012/TI-D
    If Valida = False Then Exit Sub
    '***Fin Agregado por ELRO*************************************
    
    lbGrabaNegocio = False
    
    lsNroDoc = ""
    lsDocNotaAC = ""
    lsDocumento = ""
    lsNroNotaAC = ""
    lsNroVoucher = ""
    lnTpoDoc = -1
    
    'modicado por ande 20170901 ERS029-2017
    If MsgBox(cMensaje, vbYesNo + vbQuestion, "Aviso") = vbYes Then
        cmdAceptar.Caption = "Guardando..."
        cmdAceptar.Enabled = False
        cmdCancelar.Enabled = False
        For i = 0 To nRows
        'inicio
            'lnCtaPlantillaNeg = lswPendBanc.ListItems.Item(lswPendBanc.SelectedItem.Index).SubItems(6)
            lnCtaPlantillaNeg = aRegistros(i, 6)
            lnCtaPlantillaNeg = Mid(lnCtaPlantillaNeg, 1, 2) & Mid(gsOpeCod, 3, 1) & Mid(lnCtaPlantillaNeg, 4, Len(Trim(lnCtaPlantillaNeg)))
            lnCtaCompara = lnCtaPlantillaNeg
                    
            If Mid(lnCtaPlantillaNeg, 1, 10) = "2918070101" Or Mid(lnCtaPlantillaNeg, 1, 10) = "2928070101" Then
                'lsAgeCod = lswPendBanc.ListItems.Item(lswPendBanc.SelectedItem.Index).SubItems(7)
                lsAgeCod = aRegistros(i, 7)
                lnCtaPlantillaNeg = Mid(lnCtaPlantillaNeg, 1, Len(Trim(lnCtaPlantillaNeg)) - 2) & lsAgeCod
                lnCtaCompara = Mid(lnCtaPlantillaNeg, 1, Len(Trim(lnCtaPlantillaNeg)) - 2)
            End If
                    
            If Mid(lnCtaPlantillaNeg, 1, 4) = "1116" Or Mid(lnCtaPlantillaNeg, 1, 4) = "1126" Then
                'lsAgeCod = lswPendBanc.ListItems.Item(lswPendBanc.SelectedItem.Index).SubItems(7)
                lsAgeCod = aRegistros(i, 7)
                lnCtaPlantillaNeg = Mid(lnCtaPlantillaNeg, 1, Len(Trim(lnCtaPlantillaNeg)) - 2) & lsAgeCod
                lnCtaCompara = Mid(lnCtaPlantillaNeg, 1, Len(Trim(lnCtaPlantillaNeg)) - 2)
            End If
                    
            If lbDeposito Then
                Set R2 = oOpe.CargaOpeCtaPlantillaNeg(lnCtaCompara, gsOpeCod, "H", "1")
                If Not R2.BOF And Not R2.EOF Then
                    lsCtaPendiente = lnCtaPlantillaNeg
                Else
                    MsgBox "Cuenta Contable Pendiente No esta definida", vbInformation, "AVISO"
                    'Exit Sub
                    GoTo Salir
                End If
            Else
                Set R2 = oOpe.CargaOpeCtaPlantillaNeg(lnCtaCompara, gsOpeCod, "D", "1")
                If Not R2.BOF And Not R2.EOF Then
                    lsCtaPendiente = lnCtaPlantillaNeg
                Else
                    MsgBox "Cuenta Contable Pendiente No esta definida", vbInformation, "AVISO"
                    'Exit Sub
                    GoTo Salir
                End If
            End If
                    
                    
            If gsOpeCod = "401583" Or "402583" Then
                'lsGlosa = "Deposito Banco" JACA 20110727
                'lsGlosa = lswPendBanc.ListItems.Item(lswPendBanc.SelectedItem.Index).SubItems(5)
                lsGlosa = aRegistros(i, 5) 'ANDE 20170829 ERS-029
            Else
                lsGlosa = "Retiro Banco"
            End If
                    
            If cboDocumento <> "" Then
                lnTpoDoc = Val(Right(cboDocumento, 2))
                cTpoDoc = cTpoDoc & CStr(lnTpoDoc) & "," 'ande 20170809 ers029-2017
                If lbDeposito Then
                    lsEntiOrig = gsNomCmac
                    lsCtaEntOrig = ""
                    lsEntiDest = lblDescbanco
                    lsCtaEntDest = Trim(lblDescCtabanco)
                Else
                    lsEntiOrig = lblDescbanco
                    lsCtaEntOrig = Trim(lblDescCtabanco)
                    lsEntiDest = gsNomCmac
                    lsCtaEntDest = ""
                End If
                    
                'lnImporte = lswPendBanc.ListItems.Item(lswPendBanc.SelectedItem.Index).SubItems(3)
                'lnMovnroRef = lswPendBanc.ListItems.Item(lswPendBanc.SelectedItem.Index).SubItems(8)
                lnImporte = aRegistros(i, 3)
                lnMovnroRef = aRegistros(i, 8)
                          
                Select Case lnTpoDoc
                    Case TpoDocCarta
                        oDoc.InicioCarta "", "", gsOpeCod, gsOpeDesc, lsGlosa, "", lnImporte, gdFecSis, lsEntiOrig, lsCtaEntOrig, lsEntiDest, lsCtaEntDest, ""
                        If oDoc.vbOk Then
                            lsNroDoc = oDoc.vsNroDoc
                            lsDocumento = oDoc.vsDocumento
                            cDocumento = cDocumento & lsDocumento & ","
                        Else
                            'Exit Sub
                            GoTo Salir
                        End If
                    Case TpoDocCheque
                        lsSubCtaIF = oCtaIf.SubCuentaIF(Mid(txtBuscarBanco, 4, 13))
                        lsEntiDest = gsNomCmac
                        'oDoc.InicioCheque "", True, Mid(Me.txtBuscarBanco, 4, 13), gsOpeCod, lsEntiDest, gsOpeDesc, lsGlosa, lnImporte, gdFecSis, gsNomCmacRUC, lsSubCtaIF, lsEntiOrig, lsCtaEntOrig, "", True, gsCodAge, Mid(Me.txtBuscarBanco, 18, 10)
                        oDoc.InicioCheque "", True, Mid(Me.txtBuscarBanco, 4, 13), gsOpeCod, lsEntiDest, gsOpeDesc, lsGlosa, lnImporte, gdFecSis, gsNomCmacRUC, lsSubCtaIF, lsEntiOrig, lsCtaEntOrig, "", True, gsCodAge, Mid(Me.txtBuscarBanco, 18, 10), , Mid(Me.txtBuscarBanco, 1, 2), Mid(Me.txtBuscarBanco, 4, 13), Mid(Me.txtBuscarBanco, 18, 10) 'EJVG20121130
                        If oDoc.vbOk Then
                            lsNroDoc = oDoc.vsNroDoc
                            lsDocumento = oDoc.vsDocumento
                            cDocumento = cDocumento & lsDocumento & ","
                            lsNroVoucher = oDoc.vsNroVoucher
                        Else
                            'Exit Sub
                            GoTo Salir
                        End If
                    Case TpoDocOrdenPago
                        oDoc.InicioOrdenPago "", True, "", gsOpeCod, gsNomCmac, gsOpeDesc, lsGlosa, CCur(lnImporte), gdFecSis, "", True, gsCodAge
                        If oDoc.vbOk Then
                            lsNroDoc = oDoc.vsNroDoc
                            lsDocumento = oDoc.vsDocumento
                            cDocumento = cDocumento & lsDocumento & ","
                            lsNroVoucher = oDoc.vsNroVoucher
                        Else
                            'Exit Sub
                            GoTo Salir
                        End If
                    
                    Case Else
                        lsNroDoc = txtNroDoc
                End Select
            End If
                    
            If cboDocumento <> "" Then
                Select Case Val(Right(cboDocumento, 2))
                Case TpoDocCarta, TpoDocCheque
                    'cargar el formulario de nota de cargo y abono para cargos por depositos o retiros a bancos
                    If MsgBox("Desea Generar Nota de Cargo/Abono por Comisión Adicional??", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso") = vbYes Then
                        If cboComision.ListIndex = -1 Or Trim(Right(cboComision, 6)) = "" Then
                            MsgBox "Debe seleccionar Motivo del Cargo", vbInformation, "¡AViso!"
                            cboComision.SetFocus
                            'Exit Sub
                            GoTo Salir
                        End If
                    
                        lnTpoDocAux = TpoDocNotaCargo
                        lbGrabaNegocio = True
                        frmNotaCargoAbono.inicio lnTpoDocAux, 0, gdFecSis, "", gsOpeCod, False
                        If frmNotaCargoAbono.vbOk Then
                            If frmNotaCargoAbono.Monto > lnImporte Then
                                MsgBox "Importe Cargado no debe superar el total de Operación", vbInformation, "Aviso"
                                Unload frmNotaCargoAbono
                                Set frmNotaCargoAbono = Nothing
                                'Exit Sub
                                GoTo Salir
                            End If
                            lnMontoAux = frmNotaCargoAbono.Monto
                            lsNroNotaAC = frmNotaCargoAbono.NroNotaCA
                            lsDocNotaAC = frmNotaCargoAbono.Glosa
                            lsCuentaAhoAux = frmNotaCargoAbono.CuentaAhoNro
                            lsPersNombre = frmNotaCargoAbono.PersNombre
                            lsPersDireccion = frmNotaCargoAbono.PersDireccion
                            lsUbigeo = frmNotaCargoAbono.PersUbigeo
                    
                            lnMotivoNACAux = frmNotaCargoAbono.Motivo
                            lsObjetoPadreAux = frmNotaCargoAbono.ObjetoMotivoPadre
                            lsObjetoCodAux = frmNotaCargoAbono.ObjetoMotivo
                    
                            'lsNroNotaAC = oDocRec.GetNroNotaCargoAbono(TpoDocNotaCargo)
                            'lsDocNotaAC = oContImp.ImprimeNotaCargoAbono(lsNroNotaAC, txtMovDesc, CCur(frmNotaCargoAbono.Monto), _
                                                        lsPersNombre, lsPersDireccion, lsUbigeo, gdFecSis, Mid(gsOpeCod, 3, 1), lsCuentaAhoAux, TpoDocNotaCargo, gsNomAge, gsCodUser)
                            lsDocNotaAC = oContImp.ImprimeNotaAbono(Format(gdFecSis, gsFormatoFechaView), CCur(frmNotaCargoAbono.Monto), lsGlosa, lsCuentaAhoAux, lsPersNombre, Trim(Left(cboDocumento, Len(cboDocumento) - 3)) & " " & lsNroDoc)
                            cDocNotaAC = cDocNotaAC & lsDocNotaAC & ","
                            
                            Dim oDis As New NRHProcesosCierre
                            lsCadBol = oDis.ImprimeBoletaCad(CDate(gdFecSis), "CARGO CAJA GENERAL", "CARGO CAJA GENERAL*Nro." & lsNroNotaAC, "", CCur(frmNotaCargoAbono.Monto), lsPersNombre, lsCuentaAho, "", 0, 0, "Nota Cargo", 0, 0, False, False, , , , True, , , , False, gsNomAge) & oImpresora.gPrnSaltoPagina
                    
                            Unload frmNotaCargoAbono
                            Set frmNotaCargoAbono = Nothing
                        Else
                            Unload frmNotaCargoAbono
                            Set frmNotaCargoAbono = Nothing
                            'Exit Sub
                            GoTo Salir
                        End If
                    End If
                End Select
            End If
            
            Set oCont = New NContFunciones
                    
            lsMovNro = oCont.GeneraMovNro(txtFecha, gsCodAge, gsCodUser)
            cMovNro_largo = cMovNro_largo & lsMovNro & ","
                lsCtaContBanco = oOpe.EmiteOpeCta(gsOpeCod, IIf(lbDeposito, "D", "H"), , txtBuscarBanco, CtaOBjFiltroIF)
                If lsCtaContBanco = "" Then
                    MsgBox "Institución Financiera no relacionado con Cuenta de Capital", vbInformation, "¡Aviso!"
                    'Exit Sub
                    GoTo Salir
                End If
                lsAjusteCta = ""
                If lnMontoDif <> 0 Then
                    lsAjusteCta = oOpe.EmiteOpeCta(gsOpeCod, IIf((lbDeposito And lnMontoDif > 0) Or (Not lbDeposito And lnMontoDif < 0), "H", "D"), "3")
                    If lsAjusteCta = "" Then
                        MsgBox "No se definio Cuenta para realizar Operacion con Ajuste en el Orden 3", vbInformation, "¡Aviso!"
                        'Exit Sub
                        GoTo Salir
                    End If
                End If
                    
                cmdAceptar.Enabled = False
                cmdCancelar.Enabled = False
                oCaja.GrabaDeposiBancoNegocioFinan lsMovNro, gsOpeCod, lsGlosa, lsCtaContBanco, _
                                lsCtaPendiente, CCur(lnImporte), lnTpoDoc, lsNroDoc, txtFechaDoc, lsNroVoucher, _
                                 ObjEntidadesFinancieras, txtBuscarBanco, lbDeposito, lnMotivoNAC, lsObjetoPadre, lsObjetoCod, lsCuentaAho, _
                                lnTpoDocAux, lsNroNotaAC, lnMontoAux, lnMotivoNACAux, lsObjetoPadreAux, lsObjetoCodAux, _
                                lsCuentaAhoAux, lbGrabaNegocio, gbBitCentral, Right(cboComision, 6), lsAjusteCta, lnMontoDif, nVal(lblTCFijo), lnMovnroRef
            
            Next i
        End If
        cMovNro_largo = Left(cMovNro_largo, Len(cMovNro_largo) - 1)
        If cTpoDoc <> "" Then
            cTpoDoc = Left(cTpoDoc, Len(cTpoDoc) - 1)
        End If
        If cDocumento <> "" Then
            cDocumento = Left(cDocumento, Len(cDocumento) - 1)
        End If
        If cDocNotaAC <> "" Then
            cDocNotaAC = Left(cDocNotaAC, Len(cDocNotaAC) - 1)
        End If
        
        If nRows > 0 Then
            ImprimeAsientoContable cMovNro_largo, , cTpoDoc, cDocumento, , , , , , , , , , , , cDocNotaAC, , , , , , True
        Else
            ImprimeAsientoContable cMovNro_largo, , cTpoDoc, cDocumento, , , , , , , , , , , , cDocNotaAC, , , , , , False
        End If
        
               'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & "Grabó Operación "
        Set objPista = Nothing
        '****
        'ande 20171012
        cmdAceptar.Caption = "Aceptar"
        cmdAceptar.Enabled = True
        cmdCancelar.Enabled = True
        If MsgBox("Desea ingresar otra operación?", vbQuestion + vbYesNo, "¡Confirmacion!") = vbNo Then
            Unload Me
        Else
        End If
    'end ande
End If

'***Fin Modificado por ELRO el 20121002
'lswPendBanc.ListItems.Clear

'ande 20170908 ERS029-2017
Call LlenarGrid
GoTo Salir
Salir:
    cmdAceptar.Caption = "Aceptar"
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
'and modicado 20170908 ERS029-2017
Exit Sub
AceptarErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdExcel_Click()
Dim oCtaIf As NCajaCtaIF
Dim lsMoneda As String
Dim rs As ADODB.Recordset

Dim fs              As Scripting.FileSystemObject
Dim xlAplicacion    As Excel.Application
Dim xlLibro         As Excel.Workbook
Dim xlHoja1         As Excel.Worksheet
Dim lbExisteHoja    As Boolean
Dim liLineas        As Integer
Dim i               As Integer
Dim glsArchivo      As String
Dim lsNomHoja       As String

Set rs = New ADODB.Recordset

On Error GoTo ReporteExcelBancosErr
    
    glsArchivo = "Operaciones_Bancos_Negocio" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLSX"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape

            lbExisteHoja = False
            lsNomHoja = "Operaciones Bancos"
            For Each xlHoja1 In xlLibro.Worksheets
                If xlHoja1.Name = lsNomHoja Then
                    xlHoja1.Activate
                    lbExisteHoja = True
                    Exit For
                End If
            Next
            If lbExisteHoja = False Then
                Set xlHoja1 = xlLibro.Worksheets.Add
                xlHoja1.Name = lsNomHoja
            End If

            xlAplicacion.Range("A1:A1").ColumnWidth = 15
            xlAplicacion.Range("B1:B1").ColumnWidth = 37
            xlAplicacion.Range("c1:c1").ColumnWidth = 20
            xlAplicacion.Range("D1:D1").ColumnWidth = 15
            xlAplicacion.Range("E1:E1").ColumnWidth = 15
            xlAplicacion.Range("F1:F1").ColumnWidth = 30
           
            xlAplicacion.Range("A1:Z100").Font.Size = 9
            xlAplicacion.Range("A1:Z100").Font.Name = "Century Gothic"
       
            xlHoja1.Cells(1, 1) = gsNomCmac
            xlHoja1.Cells(2, 1) = "Operaciones Bancos" & Space(5) & "DEL " & Space(5) & Format(txtFecha, "dd/mm/yyyy")
            
            xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 3)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(1, 2)).Merge True
                 
                      
            liLineas = 4
            
            xlHoja1.Cells(liLineas, 1) = "Fecha"
            xlHoja1.Cells(liLineas, 2) = "Operación"
            xlHoja1.Cells(liLineas, 3) = "Banco"
            xlHoja1.Cells(liLineas, 4) = "Monto"
            xlHoja1.Cells(liLineas, 5) = "Agencia"
            xlHoja1.Cells(liLineas, 6) = "Glosa"
   
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas + 3, 1)).Font.Bold = True
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 6)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 6)).VerticalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas + 3, 1)).Merge True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 6)).EntireRow.AutoFit
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 6)).WrapText = True
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 6)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 6)).Borders.LineStyle = 1
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 6)).Interior.ColorIndex = 36 '.Color = RGB(159, 206, 238)
            
         
            
            liLineas = liLineas + 1
            Dim lsAgeCod As String
            
            If chktodas.value = 1 Then
               lsAgeCod = ""
               
            Else
                lsAgeCod = Mid(txtBuscarArea, 4, 2)
            End If

            Set rs = oDocRec.CargaPendientesNegocioBancos(Mid(gsOpeCod, 3, 1), lbDeposito, lsAgeCod)
            
         Do Until rs.EOF
         
            xlHoja1.Cells(liLineas, 1) = rs(1)
            xlHoja1.Cells(liLineas, 2) = rs(3)
            xlHoja1.Cells(liLineas, 3) = rs(7)
            xlHoja1.Cells(liLineas, 4) = rs(4)
            xlHoja1.Cells(liLineas, 5) = rs(12)
            xlHoja1.Cells(liLineas, 6) = rs(13)
            
     
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 3), xlHoja1.Cells(liLineas, 5)).Style = "Comma"
            xlHoja1.Range(xlHoja1.Cells(liLineas, 8), xlHoja1.Cells(liLineas, 8)).Style = "Comma"
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 3), xlHoja1.Cells(liLineas, 5)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 7), xlHoja1.Cells(liLineas, 7)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 11)).HorizontalAlignment = xlCenter
            
            xlHoja1.Range(xlHoja1.Cells(liLineas, 6), xlHoja1.Cells(liLineas, 6)).HorizontalAlignment = xlRight
            xlHoja1.Range(xlHoja1.Cells(liLineas, 10), xlHoja1.Cells(liLineas, 10)).HorizontalAlignment = xlRight
            
            liLineas = liLineas + 1
            rs.MoveNext
        Loop

        ExcelCuadro xlHoja1, 1, 4, 12, liLineas - 1
        
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
    
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        Call CargaArchivo(glsArchivo, App.path & "\SPOOLER\")
          
Set oCtaIf = Nothing

        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Genero Excel "
        Set objPista = Nothing
        '****
    Exit Sub
ReporteExcelBancosErr:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub

End Sub

Private Sub cmdProcesar_Click()
If chktodas.value = 0 And txtBuscarArea.Text = "" Then
   MsgBox "Debe Elegir una Agencia", vbInformation, "Aviso"
   Exit Sub
End If

Call LlenarGrid 'ande 20170902 ERS029-2017
End Sub
Public Sub LimpiarFlexPenBanc()
    Dim nTotalRow, iRow As Integer
    nTotalRow = flxPenBanc.Rows
    For iRow = 1 To nTotalRow
        flxPenBanc.EliminaFila (nTotalRow - 1)
        nTotalRow = flxPenBanc.Rows
    Next iRow
End Sub


Sub CargaPendientesNegocio(ByVal psAgeCod As String)
Dim rs As ADODB.Recordset
Dim HayDatos  As Boolean
Set rs = New ADODB.Recordset
Dim lst As ListItem
Set rs = oDocRec.CargaPendientesNegocioBancos(Mid(gsOpeCod, 3, 1), lbDeposito, psAgeCod, chkMonto.value, Val(txtMonto.Text))
Dim iRow As Integer
iRow = 1
HayDatos = False
If Not rs.EOF And Not rs.BOF Then
   'lswPendBanc.ListItems.Clear
   HayDatos = True
   rs.MoveFirst
   Do While Not rs.EOF
'      If Not rs.BOF Then
'         Set lst = lswPendBanc.ListItems.Add(, , rs(1)) 'Fecha
'          lst.SubItems(1) = rs(3) 'Operacion nombre
'          lst.SubItems(2) = rs(7) 'Banco
'          lst.SubItems(3) = Format(rs(4), "0.00") 'Monto
'          lst.SubItems(4) = rs(12) 'Agencia
'          lst.SubItems(5) = rs(13) 'Glosa
'          lst.SubItems(6) = rs(14) 'CtaCont
'          lst.SubItems(7) = rs(15) 'Agencia
'          lst.SubItems(8) = rs(0) 'nMovNRo
'          '***Modificado por ELRO el 20121001, según OYP-RFC106-2012
'          lst.SubItems(9) = rs(2) 'cOpeCod
'          lst.SubItems(10) = rs(6) 'cReferencia
'          lst.SubItems(11) = rs(5) 'cNroDoc
'          lst.SubItems(12) = rs(11) 'dDocFecha
'          '***Fin Modificado por ELRO el 20121001*******************
'     End If
        flxPenBanc.AdicionaFila , , True
        flxPenBanc.TextMatrix(iRow, 2) = rs!Fecha 'Fecha
        flxPenBanc.TextMatrix(iRow, 3) = Trim(rs!cOpeDesc) 'nombre de la operación
        flxPenBanc.TextMatrix(iRow, 4) = Trim(rs!cPersNombre) 'banco
        flxPenBanc.TextMatrix(iRow, 5) = Format(rs!nMovImporte, "0.00") 'monto
        flxPenBanc.TextMatrix(iRow, 6) = Trim(rs!Agencia) 'nombre de la agencia
        flxPenBanc.TextMatrix(iRow, 7) = Trim(rs!cMovDesc) 'glosa
        flxPenBanc.TextMatrix(iRow, 8) = rs!cCtaContCod 'cuenta contable
        flxPenBanc.TextMatrix(iRow, 9) = rs!cAgeCod 'codigo de agencia
        flxPenBanc.TextMatrix(iRow, 10) = rs!nMovNro 'N° de Mov
        flxPenBanc.TextMatrix(iRow, 11) = rs!cOpeCod 'Cod Operación
        flxPenBanc.TextMatrix(iRow, 12) = rs!cReferencia 'referencia
        flxPenBanc.TextMatrix(iRow, 13) = rs!cNroDoc 'N° de documento
        flxPenBanc.TextMatrix(iRow, 14) = rs!dDocFecha 'Fecha de documento
        iRow = iRow + 1
        rs.MoveNext
    Loop
End If

'RIRO20131212 ERS137
If gsOpeCod = "401584" Or gsOpeCod = "402584" Then
    Set rs = Nothing
    Set rs = oDocRec.CargaPendientesTransf(psAgeCod, Mid(gsOpeCod, 3, 1))
    If Not rs Is Nothing Then
        If Not rs.BOF Then
            HayDatos = True
            Do While Not rs.EOF
           
'               Set lst = lswPendBanc.ListItems.Add(, , rs(1)) 'Fecha
'               lst.SubItems(1) = rs(3) 'Operacion nombre
'               lst.SubItems(2) = rs(7) 'Banco
'               lst.SubItems(3) = Format(rs(4), "0.00") 'Monto
'               lst.SubItems(4) = rs(12) 'Agencia
'               lst.SubItems(5) = rs(13) 'Glosa
'               lst.SubItems(6) = rs(14) 'CtaCont
'               lst.SubItems(7) = rs(15) 'Agencia
'               lst.SubItems(8) = rs(0) 'nMovNRo
'               lst.SubItems(9) = rs(2) 'cOpeCod
'               lst.SubItems(10) = rs(6) 'cReferencia
'               lst.SubItems(11) = rs(5) 'cNroDoc
'               lst.SubItems(12) = rs(11) 'dDocFecha
               flxPenBanc.AdicionaFila , , True
               flxPenBanc.TextMatrix(iRow, 2) = rs!Fecha 'Fecha
               flxPenBanc.TextMatrix(iRow, 3) = Trim(rs!cOpeDesc) 'nombre de la operación
               flxPenBanc.TextMatrix(iRow, 4) = Trim(rs!cPersNombre) 'banco
               flxPenBanc.TextMatrix(iRow, 5) = Format(rs!nMovImporte, "0.00") 'monto
               flxPenBanc.TextMatrix(iRow, 6) = Trim(rs!Agencia) 'nombre de la agencia
               flxPenBanc.TextMatrix(iRow, 7) = Trim(rs!cMovDesc) 'glosa
               flxPenBanc.TextMatrix(iRow, 8) = rs!cCtaContCod 'cuenta contable
               flxPenBanc.TextMatrix(iRow, 9) = rs!cAgeCod 'codigo de agencia
               flxPenBanc.TextMatrix(iRow, 10) = rs!nMovNro 'N° de Mov
               flxPenBanc.TextMatrix(iRow, 11) = rs!cOpeCod 'Cod Operación
               flxPenBanc.TextMatrix(iRow, 12) = rs!cReferencia 'referencia
               flxPenBanc.TextMatrix(iRow, 13) = rs!cNroDoc 'N° de documento
               flxPenBanc.TextMatrix(iRow, 14) = rs!dDocFecha 'Fecha de documento
               iRow = iRow + 1
               rs.MoveNext
        Loop
        End If
    End If
End If

If HayDatos = False Then
    MsgBox "No hay datos disponible.", vbOKOnly + vbInformation, "Aviso"
End If
rs.Close
Set rs = Nothing

End Sub

Private Sub cmdRegulariza_Click()
    Dim lnNroMovAnt As Long 'numero de la operacion en el negocio
    Dim lsMovNro As String
    Dim lsGlosa As String
    Dim lnMonto As Currency
    Dim oCont As NContFunciones
    Set oCont = New NContFunciones
    
    'ande 20171012 correcion de eliminación de filas
    Dim cMensaje As String
    Dim nCantSeleccionados As Integer, i As Integer, nFilas As Integer, iSeleccionado As Integer
    Dim aIndicesCheck() As Integer
    nCantSeleccionados = 0
    nFilas = flxPenBanc.Rows
    'obteniendo el total de seleccionado
    For i = 0 To nFilas - 1
        If flxPenBanc.TextMatrix(i, 1) = "." Then
            nCantSeleccionados = nCantSeleccionados + 1
        End If
    Next i
    'redimensionando matris para los seleccionado
    ReDim aIndicesCheck(nCantSeleccionados)
    iSeleccionado = 0
    'obteniendo los seleccionados a la matris
    For i = 0 To nFilas - 1
        If flxPenBanc.TextMatrix(i, 1) = "." Then
            aIndicesCheck(iSeleccionado) = i
            iSeleccionado = iSeleccionado + 1
        End If
    Next i
    
    If nCantSeleccionados > 1 Then
        cMensaje = "¿Desea regularizar las operaciones seleccionadas?"
    Else
        cMensaje = "¿Desea regularizar la operación seleccionada?"
    End If
    'end ande
    If MsgBox(cMensaje, vbYesNo + vbQuestion, "Aviso") = vbYes Then
        'modificado por ande 20170914 ers029-2017
        cmdRegulariza.Caption = "Regularizando..."
        cmdRegulariza.Enabled = False
        lsGlosa = "Regularización de Operaciones con Bancos"
        
        'lnNroMovAnt = lswPendBanc.ListItems.Item(lswPendBanc.SelectedItem.Index).SubItems(8)
        'lnMonto = lswPendBanc.ListItems.Item(lswPendBanc.SelectedItem.Index).SubItems(3)
        i = nCantSeleccionados - 1
        Do While (i >= 0)
    '        If flxPenBanc.TextMatrix(i, 1) = "." Then
                lnNroMovAnt = flxPenBanc.TextMatrix(aIndicesCheck(i), 10)
                lnMonto = flxPenBanc.TextMatrix(aIndicesCheck(i), 5)
        
                lsMovNro = oCont.GeneraMovNro(txtFecha, gsCodAge, gsCodUser) 'nuevo numero
                oCaja.GrabaDepositoRegularizacion lsMovNro, lnNroMovAnt, gsOpeCod, lsGlosa, lnMonto
                
                flxPenBanc.EliminaFila aIndicesCheck(i)
    '        End If
            i = i - 1
        Loop
        'ande 20171012
        cmdRegulariza.Caption = "Regularizar"
        cmdRegulariza.Enabled = True
        'Call LlenarGrid
        'end ande
    End If
End Sub






Private Sub flxPenBanc_DblClick()

Dim lvItem As ListItem
Dim i As Integer
Dim Row As Integer

Row = flxPenBanc.Row
    
If flxPenBanc.Rows > 0 Then
    'fnMovNroCap = lvItem.SubItems(8)
    fnMovNroCap = flxPenBanc.TextMatrix(Row, 10)
    'fcOpeCodCap = lvItem.SubItems(9)
    fcOpeCodCap = flxPenBanc.TextMatrix(Row, 11)
    'fcRefenciaCap = lvItem.SubItems(10)
    fcRefenciaCap = flxPenBanc.TextMatrix(Row, 12)
    'fcNroDocCap = lvItem.SubItems(11)
    fcNroDocCap = flxPenBanc.TextMatrix(Row, 13)
    'fcDocFecha = lvItem.SubItems(12)
    fcDocFecha = flxPenBanc.TextMatrix(Row, 14)
    
    If fcOpeCodCap = "200103" Or fcOpeCodCap = "200203" Or fcOpeCodCap = "210103" Or _
       fcOpeCodCap = "210807" Or fcOpeCodCap = "220103" Or fcOpeCodCap = "220203" Then
       txtBuscarBanco = fcRefenciaCap
       txtBuscarBanco_EmiteDatos
       
       For i = 0 To cboDocumento.ListCount - 1
        cboDocumento.ListIndex = i
        If Trim(Right(cboDocumento, 4)) = "55" Then
            Exit For
        End If
       Next i
       
       txtNroDoc = fcNroDocCap
       txtFechaDoc = fcDocFecha
       frmBanco.Enabled = False
       fradoc.Enabled = False
       Frame1.Enabled = False
       chkRegulariza.Enabled = False
       
    Else
        txtBuscarBanco = ""
        lblDescbanco = ""
        lblDescCtabanco = ""
        cboDocumento.ListIndex = -1
        txtNroDoc = ""
        txtFechaDoc = gdFecSis
        frmBanco.Enabled = True
        fradoc.Enabled = True
        Frame1.Enabled = True
        chkRegulariza.Enabled = True
        fnMovNroCap = 0
        fcOpeCodCap = ""
        fcRefenciaCap = ""
        fcNroDocCap = ""
        fcDocFecha = ""
    End If
    
End If
End Sub


Private Sub flxPenBanc_KeyPress(KeyAscii As Integer)
    Clipboard.Clear
    If KeyAscii = 3 Then
        'copiar solo la colunma monto
        Dim iRow, iCol As Integer
        iRow = flxPenBanc.Row
        iCol = flxPenBanc.Col
        If InStr(1, "234567", CStr(iCol)) Then
            Clipboard.SetText flxPenBanc.TextMatrix(iRow, iCol), vbCFText
        End If
    End If
End Sub

Private Sub Form_Load()
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim lsCtaPendiente As String

Set oOpe = New DOperacion
Dim R1 As ADODB.Recordset
Dim R2 As ADODB.Recordset
Set R1 = New ADODB.Recordset
Set R2 = New ADODB.Recordset
Set oCtaIf = New NCajaCtaIF
Set oCaja = New nCajaGeneral

Me.Caption = Trim(Mid(gsOpeDesc, InStr(1, gsOpeDesc, ":") + 1, Len(gsOpeDesc) - InStr(1, gsOpeDesc, ":")))
txtFecha = gdFecSis
txtFechaDoc = gdFecSis

CentraForm Me
Set oDocRec = New NDocRec
txtBuscarArea.rs = oOpe.GetOpeObj(gsOpeCod, "1")
txtBuscarBanco.psRaiz = "Cuentas de Bancos"
txtBuscarBanco.rs = oOpe.GetOpeObj(gsOpeCod, "2")

FraTipoCambio.Visible = False
If Mid(gsOpeCod, 3, 1) = gMonedaExtranjera Then
   lblTCFijo = Format(gnTipCambio, "##,###,##0.000")
   FraTipoCambio.Visible = True
End If

Set R1 = oOpe.CargaOpeDoc(gsOpeCod)


Do While Not R1.EOF
   cboDocumento.AddItem R1!cDocDesc & Space(250) & R1!nDocTpo
   R1.MoveNext
Loop
cboDocumento.AddItem "NINGUNO" & Space(250)

'If lbDeposito Then
'    Set R2 = oOpe.CargaOpeCta(gsOpeCod, "H", "1")
'    lsCtaPendiente = R2!cCtaContCod
'Else
'    Set R2 = oOpe.CargaOpeCta(gsOpeCod, "D", "1")
'    lsCtaPendiente = R2!cCtaContCod
'End If

Set oOpe = Nothing
R1.Close

Set R1 = Nothing

End Sub

'Private Sub lswPendBanc_Click()
'
'Dim lvItem As ListItem
'Dim i As Integer
'
'If lswPendBanc.ListItems.Count > 0 Then
'    Set lvItem = lswPendBanc.SelectedItem
'    fnMovNroCap = lvItem.SubItems(8)
'    fcOpeCodCap = lvItem.SubItems(9)
'    fcRefenciaCap = lvItem.SubItems(10)
'    fcNroDocCap = lvItem.SubItems(11)
'    fcDocFecha = lvItem.SubItems(12)
'    If fcOpeCodCap = "200103" Or fcOpeCodCap = "200203" Or fcOpeCodCap = "210103" Or _
'       fcOpeCodCap = "210807" Or fcOpeCodCap = "220103" Or fcOpeCodCap = "220203" Then
'       txtBuscarBanco = fcRefenciaCap
'       txtBuscarBanco_EmiteDatos
'
'       For i = 0 To cboDocumento.ListCount - 1
'        cboDocumento.ListIndex = i
'        If Trim(Right(cboDocumento, 4)) = "55" Then
'            Exit For
'        End If
'       Next i
'
'       txtNroDoc = fcNroDocCap
'       txtFechaDoc = fcDocFecha
'       frmBanco.Enabled = False
'       fraDoc.Enabled = False
'       Frame1.Enabled = False
'       chkRegulariza.Enabled = False
'
'    Else
'        txtBuscarBanco = ""
'        lblDescBanco = ""
'        lblDescCtaBanco = ""
'        cboDocumento.ListIndex = -1
'        txtNroDoc = ""
'        txtFechaDoc = gdFecSis
'        frmBanco.Enabled = True
'        fraDoc.Enabled = True
'        Frame1.Enabled = True
'        chkRegulariza.Enabled = True
'        fnMovNroCap = 0
'        fcOpeCodCap = ""
'        fcRefenciaCap = ""
'        fcNroDocCap = ""
'        fcDocFecha = ""
'    End If
'
'End If
'End Sub

Private Sub txtBuscarArea_EmiteDatos()
lblAreaDesc = txtBuscarArea.psDescripcion
    If lblAreaDesc <> "" Then
        txtFecha.SetFocus
    End If
End Sub

Private Sub txtBuscarBanco_EmiteDatos()
Dim oCtaIf As NCajaCtaIF
Dim lsCtaContBanco As String
Dim oOpe As DOperacion
Set oOpe = New DOperacion
Set oCtaIf = New NCajaCtaIF

lblDescbanco = oCtaIf.NombreIF(Mid(txtBuscarBanco, 4, 13))
lblDescCtabanco = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscarBanco, 18, 10)) + " " + txtBuscarBanco.psDescripcion
lsCtaContBanco = oOpe.EmiteOpeCta(gsOpeCod, IIf(lbDeposito, "D", "H"), , txtBuscarBanco, CtaOBjFiltroIF)
    If lsCtaContBanco = "" Then
        MsgBox "Institución Financiera no tiene definida Cuenta Contable", vbInformation, "Aviso"
    End If

'***Modificado por ELRO el 20121001, según OYP-RFC106-2012
If fcOpeCodCap = "200103" Or fcOpeCodCap = "200203" Or fcOpeCodCap = "210103" Or _
   fcOpeCodCap = "210807" Or fcOpeCodCap = "220103" Or fcOpeCodCap = "220203" Then
    cmdAceptar.SetFocus
Else
    If txtBuscarBanco <> "" Then
        cboDocumento.SetFocus
    End If
End If
'***Fin Modificado por ELRO el 20121001*******************
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdProcesar.SetFocus
End If
End Sub

'***Agregado por ELRO el 20120112, según Acta N° 003-2012/TI-D
Private Function Valida() As Boolean
Dim dFechaCierreMensualContabilidad, dFechaHabil As Date
Dim i, nDias As Integer
Dim oNConstSistemas As NConstSistemas
Set oNConstSistemas = New NConstSistemas
Dim oNContFunciones As New NContFunciones
Set oNContFunciones = New NContFunciones
   
dFechaCierreMensualContabilidad = CDate(oNConstSistemas.LeeConstSistema(gConstSistCierreMensualCont))

Valida = True

If Month(CDate(txtFecha)) = Month(dFechaCierreMensualContabilidad) And _
   Year(CDate(txtFecha)) = Year(dFechaCierreMensualContabilidad) Then
    
    If MsgBox("¿Desea realizar la operación en una fecha que pertenece a un Mes Cerrado?", vbYesNo, "Confirmar") = vbYes Then
    nDias = DateDiff("D", dFechaCierreMensualContabilidad, gdFecSis)
        For i = 1 To nDias
        
            If Not oNContFunciones.EsFeriado(DateAdd("D", i, dFechaCierreMensualContabilidad)) Then
                dFechaHabil = DateAdd("D", i, dFechaCierreMensualContabilidad)
                If DateDiff("D", dFechaHabil, gdFecSis) > 0 Then
                    MsgBox "Solo se puede realizar la operación en un Mes Cerrado hasta " & dFechaHabil, vbInformation, "aviso"
                    Valida = False
                    txtFecha.SetFocus
                    Exit Function
                    
                End If
            
            End If
         
        Next i
    Else
        Valida = False
        Exit Function
    End If
Else
    If Not ValidaFechaContab(txtFecha, gdFecSis, True) Then
       Valida = False
       fEnfoque txtFecha
       txtFecha.SetFocus
       Exit Function
    End If
End If

End Function

'***Fin Agregado por ELRO*************************************

