VERSION 5.00
Begin VB.Form frmSegListaGarant 
   Caption         =   "Lista - Garantias"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraCuentas 
      Caption         =   "Garantias"
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
      Height          =   2505
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      Begin SICMACT.FlexEdit feGarantias 
         Height          =   2175
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3836
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Garantia-Suma Aseguada-cNumGarant-Moneda"
         EncabezadosAnchos=   "400-3400-1250-0-700"
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
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-C-C"
         FormatosEdit    =   "0-0-2-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1635
      TabIndex        =   1
      Top             =   2580
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2775
      TabIndex        =   0
      Top             =   2595
      Width           =   1000
   End
End
Attribute VB_Name = "frmSegListaGarant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vbTIpo As Integer
Dim lsNumGaran As String, lsGarantDesc As String, lsMoneda As String
Dim lnSumaAsegurada As Double
Dim lsTipoSeguro As String

Public Sub Inicia(ByVal psPersCod As String, ByRef psNumGaran As String, ByRef psGarantDesc As String, ByRef pnSumaAsegurada As Double, ByRef psMoneda As String, ByVal psTpoSeg As TipoSeguro)
    lsTipoSeguro = psTpoSeg
    Call CargarListaGarant(psPersCod)
    lsNumGaran = ""
    Me.Show 1
    psNumGaran = lsNumGaran
    psGarantDesc = lsGarantDesc
    pnSumaAsegurada = lnSumaAsegurada
    psMoneda = lsMoneda
End Sub

Private Sub CmdAceptar_Click()
    If feGarantias.TextMatrix(feGarantias.row, 3) = "" Then
        MsgBox "Debe seleccionar un dato valido.", vbExclamation, "Alerta"
    Else
        lsNumGaran = feGarantias.TextMatrix(feGarantias.row, 3)
        lsGarantDesc = feGarantias.TextMatrix(feGarantias.row, 1)
        lnSumaAsegurada = feGarantias.TextMatrix(feGarantias.row, 2)
        lsMoneda = feGarantias.TextMatrix(feGarantias.row, 4)
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Garantias de la Persona"
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Public Sub CargarListaGarant(ByVal psPersCod As String)
    Dim obj As New COMNCredito.NCOMGarantia
    Dim oDR As New Recordset
    Dim i As Integer
    
    'RECO20150715 ERS304-2015 ***************************************
    'Set oDR = obj.RecuperaListaGarantPolizaPersona(psPersCod)
    Select Case lsTipoSeguro
        Case gContraIncendio
            Set oDR = obj.RecuperaListaGarantPolizaPersona(psPersCod)
        Case 10030000
            Set oDR = obj.RecuperaListaGarantMoviliariaPersona(psPersCod)
    End Select
    'RECO FIN *******************************************************
    If Not (oDR.EOF And oDR.BOF) Then
        feGarantias.Clear
        FormateaFlex feGarantias
        For i = 1 To oDR.RecordCount
            feGarantias.AdicionaFila
            feGarantias.TextMatrix(i, 1) = oDR!cDescripcion
            feGarantias.TextMatrix(i, 2) = oDR!nSumaAsegurada
            feGarantias.TextMatrix(i, 3) = oDR!cNumGarant
            feGarantias.TextMatrix(i, 4) = oDR!cmoneda
            oDR.MoveNext
        Next
    End If
    Set oDR = Nothing
End Sub
