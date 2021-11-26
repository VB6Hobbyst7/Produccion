VERSION 5.00
Begin VB.Form frmCredSeguimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seguimiento de Proceso de Créditos"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16935
   Icon            =   "frmCredSeguimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   16935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelec 
      Caption         =   "Ver detalle Selección"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   2535
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   15480
      TabIndex        =   10
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   " Lista de Créditos "
      ForeColor       =   &H00FF0000&
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   16815
      Begin SICMACT.FlexEdit feCreditos 
         Height          =   3855
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   16575
         _ExtentX        =   29236
         _ExtentY        =   6800
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Agencia-Crédito-Titular-Producto-Monto-Moneda-Analista-Ubicación"
         EncabezadosAnchos=   "500-2100-1800-4200-2100-1200-900-1200-2400"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-L-L-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Opciones de Filtrado "
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   320
         Left            =   5280
         TabIndex        =   8
         Top             =   910
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar..."
         Height          =   320
         Left            =   5280
         TabIndex        =   7
         Top             =   610
         Width           =   1455
      End
      Begin VB.ComboBox cboAgencia 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   920
         Width           =   3735
      End
      Begin VB.TextBox txtTitular 
         Height          =   285
         Left            =   1410
         TabIndex        =   5
         Top             =   620
         Width           =   3735
      End
      Begin VB.OptionButton optAge 
         Caption         =   "Agencia"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optTit 
         Caption         =   "Titular"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   650
         Width           =   975
      End
      Begin VB.OptionButton optCredito 
         Caption         =   "Crédito"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1290
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCredSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCredSeguimiento
'*** Descripción : Formulario para realizar el seguimiento de los creditos
'*** Creación : RECO el 20161020, según ERS060-2016
'********************************************************************
Option Explicit

Dim lnFilaSelec As Integer

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If
End Sub

Private Sub cboAgencia_KeyPress(KeyAscii As Integer)
    cmdBuscar.SetFocus
End Sub

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdBuscar_Click()
    Dim Msj As String
    Msj = ValidaFiltros
    If Msj = "" Then
        Screen.MousePointer = 11
        Call CargarDatos
        Call HabilitaOpciones(0, True)
        cmdSelec.Enabled = True
        Screen.MousePointer = 0
    Else
        MsgBox Msj, vbInformation, "Alerta"
    End If
End Sub

Private Sub cmdLimpiar_Click()
    Call LimpiarFormulario
    cmdSelec.Enabled = False
End Sub

Private Sub cmdSelec_Click()
    If feCreditos.TextMatrix(feCreditos.row, 1) <> "" Then
        frmCredSeguimientoDet.Inicia (feCreditos.TextMatrix(feCreditos.row, 2))
    Else
        MsgBox "No se encontraron datos relacionados a la selección ", vbInformation, "Alerta"
    End If
End Sub

Private Sub Form_Load()
    Call CargarAgencias
    'Call LimpiarFormulario
End Sub

Private Sub CargarAgencias()
    Dim oAge As New COMDConstantes.DCOMAgencias
    Dim RS As New ADODB.Recordset
    
    Set RS = oAge.ObtieneAgencias
    Call CargaCombo(cboAgencia, RS)
    cboAgencia.ListIndex = 0
End Sub

Private Function ValidaFiltros() As String
    ValidaFiltros = ""
    If optCredito.value = True Then
        If ActXCodCta.NroCuenta = "" Then
            ValidaFiltros = "Ingrese el número de crédito"
        ElseIf Len(ActXCodCta.NroCuenta) < 18 Then
            ValidaFiltros = "Número de crédito incorrecto"
        End If
    ElseIf optTit.value = True And txtTitular.Text = "" Then
        ValidaFiltros = "Ingrese el nombre del Titular"
    End If
End Function

Private Function HabilitaOpciones(ByVal pnOpcion As Integer, Optional ByVal pbBuscar As Boolean)
    If pbBuscar = True Then
        pnOpcion = 0
    End If
    optCredito.Enabled = Not pbBuscar
    optTit.Enabled = Not pbBuscar
    optAge.Enabled = Not pbBuscar
    cmdBuscar.Enabled = Not pbBuscar
    
    ActXCodCta.Enabled = IIf(pnOpcion = 1, True, False)
    txtTitular.Enabled = IIf(pnOpcion = 2, True, False)
    cboAgencia.Enabled = IIf(pnOpcion = 3, True, False)
    If pnOpcion = 1 Then
        ActXCodCta.SetFocus
        txtTitular.Text = ""
        cboAgencia.ListIndex = 0
    ElseIf pnOpcion = 2 Then
        txtTitular.SetFocus
        cboAgencia.ListIndex = 0
        ActXCodCta.NroCuenta = ""
    ElseIf pnOpcion = 3 Then
        cboAgencia.SetFocus
        ActXCodCta.NroCuenta = ""
        txtTitular.Text = ""
    End If
End Function

Private Sub optAge_Click()
    Call HabilitaOpciones(3)
End Sub

Private Sub optAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboAgencia.SetFocus
    End If
End Sub

Private Sub optCredito_Click()
    Call HabilitaOpciones(1)
End Sub

Private Sub optCredito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ActXCodCta.SetFocus
    End If
End Sub

Private Sub optTit_Click()
    Call HabilitaOpciones(2)
End Sub

Private Sub optTit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTitular.SetFocus
    End If
End Sub

Private Sub txtTitular_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If
End Sub

Private Sub LimpiarFormulario()
    ActXCodCta.NroCuenta = ""
    txtTitular.Text = ""
    cboAgencia.ListIndex = 0
    feCreditos.Clear
    FormateaFlex feCreditos
    optCredito.value = True
    cmdSelec.Enabled = False
    Call HabilitaOpciones(1, False)
End Sub

Private Sub CargarDatos()
    If optCredito.value = True Then
        Call CargarLista(ActXCodCta.NroCuenta, "", "")
    ElseIf optTit.value = True Then
        Call CargarLista("", txtTitular.Text, "")
    Else
        Call CargarLista("", "", Mid(cboAgencia.Text, Len(cboAgencia.Text) - 1, 2))
    End If
End Sub

Private Sub txtTitular_LostFocus()
    txtTitular.Text = UCase(txtTitular.Text)
End Sub

Private Sub CargarLista(ByVal psCtaCod As String, ByVal psNombre As String, ByVal psAgencia As String)
Dim oEvalN As New COMNCredito.NCOMColocEval
    Dim RS As New ADODB.Recordset
    Dim nIndice As Integer
    
    
    
    Set RS = oEvalN.ListaCreditosSeguimiento(psCtaCod, psNombre, psAgencia)
    If Not (RS.EOF And RS.BOF) Then
        MsgBox "Se está cargando los datos, espere un momento..", vbInformation, "Aviso" 'ARLO 20170717
        feCreditos.Clear
        FormateaFlex feCreditos
        For nIndice = 1 To RS.RecordCount
            feCreditos.AdicionaFila
            feCreditos.TextMatrix(nIndice, 0) = nIndice
            feCreditos.TextMatrix(nIndice, 1) = RS!Agencia
            feCreditos.TextMatrix(nIndice, 2) = RS!Credito
            feCreditos.TextMatrix(nIndice, 3) = RS!Titular
            feCreditos.TextMatrix(nIndice, 4) = RS!Producto
            feCreditos.TextMatrix(nIndice, 5) = RS!Monto
            feCreditos.TextMatrix(nIndice, 6) = RS!Moneda
            feCreditos.TextMatrix(nIndice, 7) = RS!analista 'ARLO 20170921
            feCreditos.TextMatrix(nIndice, 8) = RS!ubicacion
            RS.MoveNext
        Next
    'End If
    Else
    MsgBox "No existen Registros", vbInformation, "Aviso" 'ARLO 20170717
    End If
End Sub
