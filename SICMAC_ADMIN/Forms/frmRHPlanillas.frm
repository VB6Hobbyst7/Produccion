VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRHPlanillas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12165
   Icon            =   "frmRHPlanillas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexportar 
      Caption         =   "Exportar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4365
      TabIndex        =   39
      Top             =   6675
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   6675
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbonarCuentas 
      Caption         =   "&Abonar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8445
      TabIndex        =   6
      Top             =   6675
      Width           =   1095
   End
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "&Asiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7365
      MaskColor       =   &H0000FFFF&
      TabIndex        =   35
      Top             =   6675
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picSi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4260
      Picture         =   "frmRHPlanillas.frx":030A
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   30
      Top             =   7200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4710
      Picture         =   "frmRHPlanillas.frx":064C
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   29
      Top             =   7200
      Visible         =   0   'False
      Width           =   315
   End
   Begin Sicmact.TxtBuscar TxtPlanillas 
      Height          =   315
      Left            =   1065
      TabIndex        =   25
      Top             =   30
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
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
   End
   Begin VB.ComboBox cmbOpc 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmRHPlanillas.frx":098E
      Left            =   8400
      List            =   "frmRHPlanillas.frx":0990
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   30
      Width           =   2190
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7365
      MaskColor       =   &H0000FFFF&
      TabIndex        =   19
      Top             =   6675
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3270
      TabIndex        =   3
      Top             =   6675
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2190
      TabIndex        =   2
      Top             =   6675
      Width           =   1095
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   1110
      TabIndex        =   1
      Top             =   6675
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   6675
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox R 
      Height          =   105
      Left            =   4695
      TabIndex        =   7
      Top             =   6735
      Visible         =   0   'False
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   185
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"frmRHPlanillas.frx":0992
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraPla 
      Appearance      =   0  'Flat
      Caption         =   "Planilla"
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
      Height          =   6270
      Left            =   15
      TabIndex        =   8
      Top             =   330
      Width           =   12120
      Begin VB.TextBox txtCambFij 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6060
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   585
         Width           =   855
      End
      Begin VB.TextBox txtTpoCambio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7875
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   585
         Width           =   855
      End
      Begin Sicmact.TxtBuscar txtPlanillaIns 
         Height          =   300
         Left            =   825
         TabIndex        =   31
         Top             =   210
         Width           =   1620
         _ExtentX        =   2858
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
         TipoBusqueda    =   2
      End
      Begin VB.TextBox txtDes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   810
         MaxLength       =   254
         TabIndex        =   10
         Top             =   975
         Width           =   7920
      End
      Begin VB.ComboBox cmbEstadoPla 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   585
         Width           =   4020
      End
      Begin MSMask.MaskEdBox txtFecFin 
         Height          =   285
         Left            =   9345
         TabIndex        =   11
         Top             =   960
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecIni 
         Height          =   285
         Left            =   9345
         TabIndex        =   12
         Top             =   600
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
         Height          =   4845
         Left            =   0
         TabIndex        =   13
         Top             =   1320
         Width           =   12105
         _ExtentX        =   21352
         _ExtentY        =   8546
         _Version        =   393216
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSMask.MaskEdBox mskFecImp 
         Height          =   285
         Left            =   9345
         TabIndex        =   27
         Top             =   210
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTipCambFF 
         Caption         =   "Tpo.Camb FF:"
         Height          =   240
         Left            =   4890
         TabIndex        =   37
         Top             =   615
         Width           =   1095
      End
      Begin VB.Label lblTpoCambio 
         Caption         =   "Tpo.Camb:"
         Height          =   240
         Left            =   7050
         TabIndex        =   33
         Top             =   615
         Width           =   825
      End
      Begin VB.Label lblPlaniInstRes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2535
         TabIndex        =   32
         Top             =   210
         Width           =   6195
      End
      Begin VB.Label lblFecImp 
         Caption         =   "Impre... :"
         Height          =   225
         Left            =   8850
         TabIndex        =   28
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Fin"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   8850
         TabIndex        =   18
         Top             =   990
         Width           =   405
      End
      Begin VB.Label lblFecIni 
         Caption         =   "Inicio"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   8850
         TabIndex        =   17
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "Descrip."
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   990
         Width           =   735
      End
      Begin VB.Label lblPlanilla 
         Caption         =   "Planilla :"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   90
         TabIndex        =   15
         Top             =   225
         Width           =   735
      End
      Begin VB.Label lblEstado 
         Caption         =   "Estado:"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   90
         TabIndex        =   14
         Top             =   630
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9600
      TabIndex        =   21
      Top             =   6675
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelarPlanilla 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1110
      TabIndex        =   22
      Top             =   6675
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   30
      TabIndex        =   20
      Top             =   6675
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdValida 
      Caption         =   "&Valida.Pla."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6285
      MaskColor       =   &H0000FFFF&
      TabIndex        =   38
      Top             =   6675
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "DOLARES:"
      Height          =   255
      Left            =   8520
      TabIndex        =   44
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "SOLES:"
      Height          =   255
      Left            =   6120
      TabIndex        =   43
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label lblSol 
      Height          =   255
      Left            =   6840
      TabIndex        =   42
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label lblDol 
      Height          =   255
      Left            =   9720
      TabIndex        =   41
      Top             =   7200
      Width           =   1095
   End
   Begin VB.OLE OLE1 
      Height          =   255
      Left            =   5760
      TabIndex        =   40
      Top             =   6720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblPlanillaRes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2580
      TabIndex        =   26
      Top             =   30
      Width           =   5295
   End
   Begin VB.Label lblTipPlanilla 
      Caption         =   "Tipo Planilla :"
      Height          =   180
      Left            =   60
      TabIndex        =   24
      Top             =   45
      Width           =   1005
   End
   Begin VB.Label lblOpciones 
      Caption         =   "Opc :"
      Height          =   195
      Left            =   7950
      TabIndex        =   23
      Top             =   75
      Width           =   405
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuBuscar 
         Caption         =   "&Buscar..."
      End
      Begin VB.Menu mnuBuscarSiguiente 
         Caption         =   "B&uscar Siguiente"
      End
      Begin VB.Menu mnuComentario 
         Caption         =   "&Comentario"
      End
      Begin VB.Menu mnuGuion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAgregar 
         Caption         =   "&Agregar"
      End
   End
End
Attribute VB_Name = "frmRHPlanillas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim lbEditado As Boolean
Dim lsCadBol As String
Dim lnColorB As Long
Dim lnColorP As Long
Dim lsPorMesTrab() As String
Dim lnSalir As Integer
Dim lbCancela As Boolean
Dim lnIdxBuscar As Integer
Dim lsCadenaBuscar As String
Dim lnUtilDias As Currency
Dim lnNumMonto As Currency
Dim lnTotSueldo As Currency
Dim lnTotDias As Currency
Dim lnTipo As TipoProcesoRRHH
Dim lsPlanillaAnt As String
Dim WithEvents oPlaEvento As NActualizaDatosConPlanilla
Attribute oPlaEvento.VB_VarHelpID = -1
Dim Progress As clsProgressBar
Dim lsPlanillaCodDefecto As String

Dim lsMov1 As String
Dim lsMov2 As String
'Para Exportar a Excel
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lsOpeCodFractal As String 'ALPA 20120417

Public Sub IniResincion(pnTipo As TipoProcesoRRHH, psCaption As String, pMdi As Form, Optional psPlanillaCodDefecto As String = "", Optional psCodPers As String = "", Optional psPersNombre As String = "", Optional psCodEmp As String = "")
    lsPlanillaCodDefecto = psPlanillaCodDefecto
    lnTipo = pnTipo
    Caption = psCaption
    CmdNuevo_Click
    
    If Flex.TextMatrix(Flex.Rows - 1, 1) <> "" Then
        Flex.Rows = Flex.Rows + 1
    End If
    
    Flex.TextMatrix(Flex.Rows - 1, 0) = "1"
    Flex.TextMatrix(Flex.Rows - 1, 1) = psCodEmp
    Flex.TextMatrix(Flex.Rows - 1, 2) = psCodPers
    Flex.TextMatrix(Flex.Rows - 1, 3) = psPersNombre
    Flex.row = Flex.Rows - 1
    Flex.Col = 4
    Set Flex.CellPicture = picSi
    
    Me.Show , pMdi
End Sub

Private Sub cmbEstadoPla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDes.SetFocus
End Sub

Private Sub cmbOpc_Change()
    Dim i As Integer
    Dim lnPos As Integer
    
    If cmbOpc.ListIndex = -1 Then Exit Sub
        
    Flex.Col = 4
    If Trim(Right(cmbOpc, 5)) = "1" Then
        If IsNumeric(Mid(Flex.TextMatrix(Flex.Rows - 1, 1), 2)) Then
            For i = 1 To Flex.Rows - 1
                Flex.row = i
                Flex.TextMatrix(i, 0) = "1"
                Set Flex.CellPicture = picSi
            Next i
        Else
            For i = 1 To Flex.Rows - 2
                Flex.row = i
                Flex.TextMatrix(i, 0) = "1"
                Set Flex.CellPicture = picSi
            Next i
        End If
    ElseIf Trim(Right(cmbOpc, 5)) = "2" Then
        If IsNumeric(Mid(Flex.TextMatrix(Flex.Rows - 1, 1), 2)) Then
            For i = 1 To Flex.Rows - 1
                Flex.row = i
                Flex.TextMatrix(i, 0) = "0"
                Set Flex.CellPicture = picNo
            Next i
        Else
            For i = 1 To Flex.Rows - 2
                Flex.row = i
                Flex.TextMatrix(i, 0) = "0"
                Set Flex.CellPicture = picNo
            Next i
        End If
    ElseIf Trim(Right(cmbOpc, 5)) = "4" Then
        For i = 1 To Flex.Rows - 2
            Flex.row = i
            If Left(Flex.TextMatrix(i, 1), 1) = "_" Or Left(Flex.TextMatrix(i, 1), 1) = "" Then
                Flex.TextMatrix(i, 0) = "0"
                Set Flex.CellPicture = picNo
            Else
                Flex.TextMatrix(i, 0) = "1"
                Set Flex.CellPicture = picSi
            End If
        Next i
    ElseIf Trim(Right(cmbOpc, 5)) = "5" Then
        GetPosIJ Flex.TextMatrix(1, 1), "I_TOT_ING", 1, lnPos
        For i = 1 To Flex.Rows - 2
            Flex.row = i
            If CCur(Flex.TextMatrix(i, lnPos)) <= 0 Then
                Flex.TextMatrix(i, 0) = "0"
                Set Flex.CellPicture = picNo
            End If
        Next i
    End If
    
End Sub

Private Sub cmbOpc_Click()
    cmbOpc_Change
End Sub

Private Sub cmdAbonarCuentas_Click()
    Dim i As Integer
    Dim lnJI As Integer
    Dim lbJD As Integer
    Dim lnJD As Integer
    Dim lsCodCta As String
    Dim lsNumAboCarS As String
    Dim lsNum As String
    Dim sqlVS As String
    Dim lsMonto As String
    Dim lnSaldo As Currency
    Dim sqlE As String
    Dim lsCodigoAgencia As String
    Dim oPla As NRHProcesosCierre
    Set oPla = New NRHProcesosCierre
    Dim rsPla As ADODB.Recordset
    Set rsPla = New ADODB.Recordset
    
    Dim rsValidaPla As ADODB.Recordset
    Set rsValidaPla = New ADODB.Recordset
    
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    Dim lsDoc As String
    
    Dim cValidaAbonoCorrecto As String '*** PEAC 20131111
    Dim lsValidaDoc As String
    
    If txtPlanillaIns.Text = "" Then
        If txtPlanillaIns.Enabled Then
            MsgBox "Debe Elegir una Planilla.", vbInformation, "Aviso"
            txtPlanillaIns.SetFocus
        Else
            MsgBox "No puede abonar desde una planilla nueva.", vbInformation, "Aviso"
            cmdProcesar.SetFocus
        End If
        Exit Sub
    End If
    
    'cmdAsiento_Click
    If TxtPlanillas = "E05" Then 'Add by GITU 16/06/2008
        If MsgBox("Revisar que el Tipo de Cambio sea el correcto " + Chr(10) + "Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        txtTpoCambio.SetFocus
    End If
        
    If MsgBox("Desea Procesar Abonar a cuentas la Planilla ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    '*** PEAC 20131111
    lsCadBol = ""
    Set rsValidaPla = FlexARecordSet(Flex)
    lsValidaDoc = oPla.ValidaCtasParaAbono(rsValidaPla, Left(Me.txtPlanillaIns.Text, 8), Me.TxtPlanillas.Text, Me.txtDes.Text, CCur(Me.txtTpoCambio.Text), gsNomAge, gsEmpresa, gdFecSis, gsCodAge, gsCodUser)
    If lsValidaDoc <> "" Then
        MsgBox "EL ABONO TIENE OBSERVACIONES.", vbOKOnly + vbExclamation, "Atención"
        oPrevio.Show lsValidaDoc, Caption, False, , gImpresora
        Exit Sub
    End If
        
    '*** FIN PEAC
    Set rsPla = FlexARecordSet(Flex)
    
    If Not gbBitCentral Then
        lsDoc = oPla.AbonoPagos(rsPla, Left(Me.txtPlanillaIns.Text, 8), Me.TxtPlanillas.Text, Me.txtDes.Text, CCur(Me.txtTpoCambio.Text), gsNomAge, gsEmpresa, gdFecSis, gsCodAge, gsCodUser, Me.lblPlaniInstRes.Caption) 'APRI 20170328 AGREGADO Me.lblPlaniInstRes.Caption
    Else
        'Abona centralizado
        lsDoc = oPla.AbonoPagos22(rsPla, Left(Me.txtPlanillaIns.Text, 8), Me.TxtPlanillas.Text, Me.txtDes.Text, CCur(Me.txtTpoCambio.Text), gsNomAge, gsEmpresa, gdFecSis, gsCodAge, gsCodUser, Me.lblPlaniInstRes.Caption) 'APRI 20170328 AGREGADO Me.lblPlaniInstRes.Caption
    End If
    
    MsgBox "EL proceso de abono ha sido finalizado.", vbInformation, "Aviso"
    
    oPrevio.Show lsDoc, Caption, False, , gImpresora
End Sub

Private Sub cmdAsiento_Click()
    On Error GoTo hErr
    Dim lsPla As String, sRemEst As String, sRemCon As String
    Dim lsAsiento As String
    Dim oPla As DRHProcesosCierre
    Set oPla = New DRHProcesosCierre
    Dim oMov As DMov
    Set oMov = New DMov
    Dim oAsi As NContImprimir
    Set oAsi = New NContImprimir
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
    Dim lnOpeCont As Long
    Dim lnOpeEst As Long
    Dim sOperacion As String
    
    Dim lsTemp As String
    Dim lcOpeDecs As String '*** PEAC 20130910
    Dim lcFecAsnto  As String
    Dim lnAsntoNvo As Integer

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim lcFechaITF As String
    
    '*** PEAC 20130906
    Do
        lcFechaITF = InputBox("Confirme fecha de declaracion de ITF.", "Fecha ITF", gdFecSis)
        If IsDate(lcFechaITF) Then
            Exit Do
        ElseIf lcFechaITF = "" Then
            Exit Sub
        End If
        MsgBox "Ingrese una fecha correcta", vbExclamation + vbOKOnly, "Atención"
    Loop
    
    lnAsntoNvo = 0 '0=no existe(nuevo) 1= ya existe
    If oPla.BuscaAsientoGenerado(lsOpeCodFractal, txtPlanillaIns, lcOpeDecs, lcFecAsnto) Then
        If MsgBox("Ya existe un asiento: " & Trim(lcOpeDecs) & " de fecha " & lcFecAsnto & " ¿Desea mostrarlo?", vbYesNo + vbQuestion, "Atención") = vbNo Then
            Exit Sub
        Else
            lnAsntoNvo = 1 ' asiento ya existe
            lcFechaITF = Mid(lcFecAsnto, 7, 2) + "/" + Mid(lcFecAsnto, 5, 2) + "/" + Mid(lcFecAsnto, 1, 4)
        End If
    Else
        If MsgBox("Desea Generar el Asiento Contable ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    '*** FIN PEAC
    
'    If MsgBox("Desea Generar el Asiento Contable ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then
'        Exit Sub
'    End If
    
    lsMov1 = ""
'    lsMov2 = ""
    
    lsPla = TxtPlanillas.Text
    
    'oPla.GetOpeContPlanilla lsPla, lnOpeEst, lnOpeCont
    lsMov1 = oPla.GetAsientoITF(mskFecImp.Text, lsOpeCodFractal, lsPla, gsCodAge, gsCodUser, CDate(lcFechaITF), Trim(txtPlanillaIns.Text) + " " + Trim(lblPlaniInstRes.Caption))
'
'    Select Case lsPla
'        Case gsRHPlanillaSueldos
'            lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaSueldosRemEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoIndeterminado, gsCodAge, gsCodUser)
'            lsMov2 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaSueldosRemCon, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoFijo, gsCodAge, gsCodUser)
'            sOperacion = gsRHPlanillaSueldosRemEst
'        Case gsRHPlanillaGratificacion
'            lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaGratificacionRemEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoIndeterminado, gsCodAge, gsCodUser)
'            lsMov2 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaGratificacionRemCon, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoFijo, gsCodAge, gsCodUser)
'            sOperacion = gsRHPlanillaGratificacionRemEst
'        Case gsRHPlanillaTercio
'            lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaTercioRemEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoIndeterminado, gsCodAge, gsCodUser)
'            lsMov2 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaTercioRemCon, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoFijo, gsCodAge, gsCodUser)
'            sOperacion = gsRHPlanillaTercioRemEst
'        Case gsRHPlanillaUtilidades
'            lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaUtilidadesRemEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoIndeterminado, gsCodAge, gsCodUser)
'            lsMov2 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaUtilidadesRemCon, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoFijo, gsCodAge, gsCodUser)
'            sOperacion = gsRHPlanillaUtilidadesRemEst
'        Case gsRHPlanillaCTS
'            lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaCTSRemEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoIndeterminado, gsCodAge, gsCodUser)
'            lsMov2 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaCTSRemCon, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoFijo, gsCodAge, gsCodUser)
'            sOperacion = gsRHPlanillaCTSRemEst
'        Case gsRHPlanillaVacaciones
'            lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaVacacionesRemEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoIndeterminado, gsCodAge, gsCodUser)
'            lsMov2 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaVacacionesRemCon, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoFijo, gsCodAge, gsCodUser)
'            sOperacion = gsRHPlanillaVacacionesRemEst
'        Case gsRHPlanillaSubsidio
'            lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaSubsidioRem, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, "", gsCodAge, gsCodUser)
'            sOperacion = gsRHPlanillaSubsidioRem
'        Case gsRHPlanillaSubsidioEnfermedad
'            lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaSubsidioEnfermedadRem, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, "", gsCodAge, gsCodUser)
'            sOperacion = gsRHPlanillaSubsidioEnfermedadRem
'        Case gsRHPlanillaBonificacionVacacinal
'            lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaBonificacionVacacinalRemEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoIndeterminado, gsCodAge, gsCodUser)
'            lsMov2 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaBonificacionVacacinalRemCon, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoFijo, gsCodAge, gsCodUser)
'            sOperacion = gsRHPlanillaBonificacionVacacinalRemEst
'        Case gsRHPlanillaBonoProductividad
'            lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaBonoProductividadRemEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoIndeterminado, gsCodAge, gsCodUser)
'            lsMov2 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaBonoProductividadRemCon, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoFijo, gsCodAge, gsCodUser)
'            sOperacion = gsRHPlanillaBonoProductividadRemEst
'        'Case gsRHPlanillaReintegro
'        '    lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaReintegroRemEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoIndeterminado, gsCodAge, gsCodUser)
'        '    lsMov2 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaReintegroRemCon, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoFijo, gsCodAge, gsCodUser)
'        Case gsRHPlanillaDev5ta
'            lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaDev5taRemEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoIndeterminado, gsCodAge, gsCodUser)
'            lsMov2 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaDev5taRemCon, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoFijo, gsCodAge, gsCodUser)
'            sOperacion = gsRHPlanillaDev5taRemEst
'        Case gsRHPlanillaMovilidad
'            lsMov1 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaMovilidadRemEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoIndeterminado, gsCodAge, gsCodUser)
'            lsMov2 = oPla.GeneraAsientoRemuneracion(gsRHPlanillaMovilidadRemCon, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoFijo, gsCodAge, gsCodUser)
'            sOperacion = gsRHPlanillaMovilidadRemEst
'        Case Else
'            If lnOpeEst = -1 Or lnOpeEst = -1 Then
'                MsgBox "No se puede generar asiento contable, porque no existen operaciones contables relacionadas con la planilla.", vbInformation, "Aviso"
'                Exit Sub
'            End If
'            If Left(lsPla, 1) = "E" Then
'                lsMov1 = oPla.GeneraAsientoRemuneracion(lnOpeEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoIndeterminado, gsCodAge, gsCodUser)
'                lsMov2 = oPla.GeneraAsientoRemuneracion(lnOpeCont, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, RHContratoTipo.RHContratoTipoFijo, gsCodAge, gsCodUser)
'                sOperacion = lnOpeEst
'            Else
'                lsMov1 = oPla.GeneraAsientoRemuneracion(lnOpeEst, lsPla, Left(Me.txtPlanillaIns.Text, 8), gdFecSis, "", gsCodAge, gsCodUser)
'                lsMov2 = ""
'                sOperacion = lnOpeEst
'                lsMov2 = lnOpeEst
'            End If
'    End Select
    
    lsAsiento = ""
    If lsMov1 <> "" And lnAsntoNvo = 0 Then
        If MsgBox("Desea Imprimir Asiento Contable Consolidado ?", vbYesNo + vbQuestion, "Generar Asiento Contable") = vbYes Then
            lsAsiento = oAsi.ImprimeAsientoContable(lsMov1, 66, 80, , , , , False) & oImpresora.gPrnSaltoPagina
            oPrevio.Show lsAsiento, Caption, True, , gImpresora
        End If
    Else
        lsAsiento = oAsi.ImprimeAsientoContable(lsMov1, 66, 80, , , , , False) & oImpresora.gPrnSaltoPagina
        oPrevio.Show lsAsiento, Caption, True, , gImpresora
    End If
    
        
'        If lsMov1 <> "" And lsMov2 <> "" Then
'            If lsMov1 = "0" Then
'               lsMov1 = lsMov2
'            End If
'            If lsMov2 = "0" Then
'               lsMov2 = lsMov1
'            End If
'            lsAsiento = oAsi.ImprimeAsientoContableConsolidado(lsMov1 + "," + lsMov2, 66, 80, , , , , False) & oImpresora.gPrnSaltoPagina
'        End If
    
'    Else
'        If lsMov1 <> "" Then
'            lsAsiento = oAsi.ImprimeAsientoContable(oMov.GetcMovNro(lsMov1), 66, 80, , , , , False) & oImpresora.gPrnSaltoPagina
'        Else
'            lsMov1 = "-1"
'        End If
'
'        If lsMov2 <> "" Then
'            lsAsiento = lsAsiento & oAsi.ImprimeAsientoContable(oMov.GetcMovNro(lsMov2), 66, 80, , , , , False)
'        Else
'            lsMov2 = "-1"
'        End If
        
'    End If
        
'    If lsMov1 = -1 And lsMov2 = -1 Then
'    Else
'        ActuializaPlanillaMovs Left(Me.txtPlanillaIns.Text, 8), lsPla, IIf(IsNumeric(lsMov1), lsMov1, 0), IIf(IsNumeric(lsMov2), lsMov2, 0)
'    End If
    
'    oPrevio.Show lsAsiento, Caption, True, , gImpresora
    
'    If MsgBox("Desea Imprimir Asiento Contable Por Agencia ?", vbYesNo + vbQuestion, "Reporte Asiento Contable por Agencia") = vbYes Then
'        'MsgBox "Hola"
'        lsTemp = ""
'        Set rs = oPla.GetAgencias
'
'        While Not rs.EOF
'            'cAgeCod cAgeDescripcion
'
'            If lsMov1 = -1 Or lsMov2 = -1 Then
'                lsMov1 = 2
'                lsMov1 = 2
'            End If
'
'
'            lsAsiento = oAsi.ImprimeAsientoContConsolPorAgencia(lsMov1 + "," + lsMov2, 66, 80, , , , , False, Left(Me.txtPlanillaIns.Text, 8), rs!cAgeCod, lsPla, sOperacion, rs!cAgeDescripcion) & oImpresora.gPrnSaltoPagina
'
'            If Len(lsAsiento) < 10 Then
'
'            Else
'                If Len(Trim(lsTemp)) = 0 Then
'                    lsTemp = lsAsiento
'                Else
'                    'lstemp = lstemp & Chr(12) & lsAsiento
'                    lsTemp = lsTemp & lsAsiento
'                End If
'            End If
'            rs.MoveNext
'        Wend
'
'        oPrevio.Show lsTemp, Caption, True, , gImpresora
'
'
'    Else
'        MsgBox "Asiento ya fue generado"
'    End If
    
    
    
Exit Sub

hErr:
    Select Case MsgBox("Error " & Err.Number & ": " & Err.Description & vbCrLf & _
                      "in Sicmact.frmRhPlanillas.cmdAsiento_Click." & vbCrLf & _
                      "The error occured at line: " & Erl, _
                      vbAbortRetryIgnore + vbCritical, "Error")
        Case vbAbort
            Screen.MousePointer = vbDefault
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Private Sub cmdCancelar_Click()
    If MsgBox("Desea Cancelar el proceso ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            lbCancela = True
    End If
End Sub

Private Sub cmdCancelarPlanilla_Click()
    ValidaBotones False
    ClearScreen False
    Activa False, False
    TxtPlanillas.Enabled = True
    
    'MAVM 20110715 ***
    If TxtPlanillas.Text = "E06" Then
        Dim objD As DActualizaDatosConPlanilla
        Set objD = New DActualizaDatosConPlanilla
        Dim i As Integer
    
        For i = 1 To Me.Flex.Rows - 2
            objD.ActualizarDiasVacaciones Flex.TextMatrix(i, 2), IIf(Flex.TextMatrix(i, 5) = "", 0, Flex.TextMatrix(i, 5)), 2, Format(Me.mskFecImp, "yyyymmdd") & gsCodUser
        Next i
    End If
    '***
End Sub

Private Sub CmdEliminar_Click()
    Dim oPla As NActualizaDatosConPlanilla
    Set oPla = New NActualizaDatosConPlanilla
    
    Dim oPle As DRHProcesosCierre
    Set oPle = New DRHProcesosCierre
    
    If Right(cmbEstadoPla, 1) = RHPlanillaEstado.RHPlanillaEstadoPagado Then
        MsgBox "No se puede Eliminar una planilla con estado PAGADO", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If TxtPlanillas.Text = "" Then
        MsgBox "Debe Elegir un tipo de Planilla.", vbInformation, "Aviso"
        TxtPlanillas.SetFocus
        Exit Sub
    End If
    
    If txtPlanillaIns.Text = "" Then
        MsgBox "Debe Elegir una Planilla.", vbInformation, "Aviso"
        txtPlanillaIns.SetFocus
        Exit Sub
    End If
    
    
    If oPle.GetContadorPago(Trim(TxtPlanillas.Text), Left(txtPlanillaIns.Text, 8)) > 0 Then
        MsgBox "La Planilla no se puede Eliminar, porque tiene Trabajadores con Pago", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Desea Eliminar la Planilla Planilla ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    
    'MAVM 20110715 ***
    If TxtPlanillas.Text = "E06" Then
        Dim objD As DActualizaDatosConPlanilla
        Set objD = New DActualizaDatosConPlanilla
        Dim i As Integer
        For i = 1 To Me.Flex.Rows - 2
            objD.ActualizarDiasVacaciones Flex.TextMatrix(i, 2), Flex.TextMatrix(i, 5), 1, Left(txtPlanillaIns.Text, 8)
        Next i
    End If
    '***
    
    oPla.EliminaPlanilla Left(txtPlanillaIns.Text, 8), Me.TxtPlanillas.Text
    Me.TxtPlanillas.Text = ""
    CargaPlanillasTpo TxtPlanillas.Text
    txtPlanillaIns_EmiteDatos
    
End Sub

Private Sub cmdExportar_Click()
    Dim i As Long
    Dim n As Long
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Dim lsCadAnt As String
    Dim lnIni As Integer
    Dim J As Integer
    Dim lsCad As String
    
    
    Dim Col As Integer
    Dim Fil As Integer
    Dim Mit As Integer
    
    Dim lsTempCod As String
    Dim lsAgenciaFinPla As String
    
    On Error Resume Next
    lsArchivoN = App.path & "\planilla.xls"
    OLE1.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If Not lbLibroOpen Then
       Err.Clear
       'Set objExcel = CreateObject("Excel.Application")
       If Err.Number Then
          MsgBox "Can't open Excel."
       End If
       Exit Sub
    End If
    Set xlHoja1 = xlLibro.Worksheets(1)
    ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
    Dim band  As Boolean
    Dim letra As String
    lnIni = 0

    Col = Flex.Cols - 1
    Fil = Flex.Rows - 1
    Mit = Col / 2

    xlAplicacion.Range("A1:A1").ColumnWidth = 7
    xlAplicacion.Range("B1:B1").ColumnWidth = 0
    xlAplicacion.Range("C1:C1").ColumnWidth = 37
    xlAplicacion.Range("D1:D1").ColumnWidth = 0

   xlHoja1.Cells(2, 1).value = "Caja Municipal de Ahorro y Credito Maynas CMAC-MAYNAS S.A. "
   xlHoja1.Cells(2, Mit).value = "PLANILLA DE: " & " " & " " & lblPlanillaRes.Caption & " " & "EMPLEADOS "
   xlHoja1.Cells(3, 3).value = "Jr. Próspero 791 RUC:20103845328"
   xlHoja1.Cells(3, Mit).value = "MES"
   xlHoja1.Cells(3, Mit + 1).value = Mid(txtFecFin.Text, 4, 2)
   xlHoja1.Cells(3, Mit + 2).value = "AÑO"
   xlHoja1.Cells(3, Mit + 3).value = Mid(txtFecIni.Text, 7, 4)

   xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, Col + 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
   xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, Col + 14)).Borders(xlEdgeBottom).Weight = xlMedium
   xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, Col + 14)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
     
   
   xlHoja1.Cells(7, Mit / 2).value = "INGRESOS"
   xlHoja1.Cells(7, Mit + 2).value = "DESCUENTOS A EMPLEADOS"
   xlHoja1.Cells(7, Col - 4).value = "DESCUENTOS AL EMPLEADOR"
   xlHoja1.Cells(7, Col + 1).value = "NETO A PAGAR"
    
    
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, Col)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(3, 1), xlHoja1.Cells(3, Col)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(7, Col + 1)).Font.Bold = True
   
   
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(2, 6)).Merge True
   xlHoja1.Range(xlHoja1.Cells(2, Mit), xlHoja1.Cells(2, Mit + 4)).Merge True
   xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 5)).Merge True
   xlHoja1.Range(xlHoja1.Cells(7, Mit + 2), xlHoja1.Cells(7, Mit + 5)).Merge True
   xlHoja1.Range(xlHoja1.Cells(7, Col - 4), xlHoja1.Cells(7, Col - 1)).Merge True
   
   xlHoja1.Range(xlHoja1.Cells(3, Mit + 1), xlHoja1.Cells(3, Mit + 1)).HorizontalAlignment = xlLeft
   xlHoja1.Range(xlHoja1.Cells(3, Mit + 3), xlHoja1.Cells(3, Mit + 3)).HorizontalAlignment = xlLeft

    Dim cont As Integer
    Dim Cont1 As Integer
    Dim Total As Currency
    
    cont = 0
    Cont1 = 0
    
    Dim oConRH As DRHConcepto
    Set oConRH = New DRHConcepto
    Dim lsTemp As String

    For i = 0 To Flex.Rows - 1
        Flex.row = i
        
        Dim rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
         
        For n = 4 To Flex.Cols - 1
            If i = 0 Then
                xlHoja1.Range(xlHoja1.Cells(9, 1), xlHoja1.Cells(9, Col + 14)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                xlHoja1.Range(xlHoja1.Cells(9, 1), xlHoja1.Cells(9, Col + 14)).Borders(xlEdgeBottom).Weight = xlMedium
                xlHoja1.Range(xlHoja1.Cells(9, 1), xlHoja1.Cells(9, Col + 14)).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
            End If
                
                Flex.Col = n
                If i = 0 Then
                    lsTemp = oConRH.GetImpreConcepto(Flex.Text)
                    If lsTemp = "" Then
                        xlHoja1.Cells(i + 9, n + 13).value = Flex.Text
                    Else
                        xlHoja1.Cells(i + 9, n + 13).value = oConRH.GetImpreConcepto(Flex.Text)
                    End If
                Else
                    xlHoja1.Cells(i + 9, n + 13).value = Flex.Text
                End If
                
                If Flex.Text = "I_TOT_ING" Then
                    cont = n
                End If
                
                If Flex.Text = "D_TOT_DESC" Then
                    Cont1 = n
                End If
                          
            If i > 0 Then
                If n = cont Then
                   Total = Abs(Val(IIf(Flex.Text = "", 0, CCur(Flex.Text))))
                Else
                  If n = Cont1 Then
                     Total = Total - Abs(Val(IIf(Flex.Text = "", 0, CCur(Flex.Text))))
                  End If
                End If
            End If
        
            '*************************************
                         
        Next
        xlHoja1.Cells(i + 9, n + 13).value = Total

        If i > 0 Then
            Flex.Col = 1
            lsTempCod = Flex.Text
            xlHoja1.Cells(i + 9, 1).value = lsTempCod
            Flex.Col = 3
            xlHoja1.Cells(i + 9, 3).value = Flex.Text
            Set rs1 = oConRH.GetDatosPersonaPlanilla(lsTempCod, Format(txtFecFin.Text, "yyyy/mm/dd"))
            If Not rs1.EOF And Not rs1.BOF Then
               xlHoja1.Cells(i + 9, 4).value = rs1!area
               xlHoja1.Cells(i + 9, 5).value = rs1!Cargo
               xlHoja1.Cells(i + 9, 6).value = rs1!Agencia
               xlHoja1.Cells(i + 9, 7).value = "'" + Format(rs1!FechaIng, "dd/mm/yyyy")
               xlHoja1.Cells(i + 9, 8).value = rs1!DNI
               xlHoja1.Cells(i + 9, 9).value = "'" + Format(rs1!FechaNaci, "dd/mm/yyyy")
               xlHoja1.Cells(i + 9, 10).value = rs1!Domicilio
               xlHoja1.Cells(i + 9, 11).value = rs1!Telefono
               'ALPA 20110118****************************************
               xlHoja1.Cells(i + 9, 12).value = IIf(IsNull(rs1!cUbicacion), rs1!Agencia, rs1!cUbicacion)
               xlHoja1.Cells(i + 9, 13).value = "'" + IIf(IsNull(rs1!dRHContratoFin), "", Format(rs1!dRHContratoFin, "dd/mm/yyyy"))
               xlHoja1.Cells(i + 9, 14).value = rs1!cConsDescripcionNivel
               xlHoja1.Cells(i + 9, 15).value = rs1!cPersNatSexo
               '*****************************************************
               xlHoja1.Cells(i + 9, 16).value = rs1!AFP
               xlHoja1.Cells(i + 9, 17).value = rs1!cCUSPP
             End If
         End If
        
    Next

    xlHoja1.Cells(9, 1).value = "Cod.Emp."
    xlHoja1.Cells(9, 3).value = "Nombre"
    xlHoja1.Cells(9, 4).value = "AREA"
    xlHoja1.Cells(9, 5).value = "CARGO"
    xlHoja1.Cells(9, 6).value = "AGENCIA"
    xlHoja1.Cells(9, 7).value = "F.Ingreso"
    xlHoja1.Cells(9, 8).value = "DNI"
    xlHoja1.Cells(9, 9).value = "FechaNacimiento"
    xlHoja1.Cells(9, 10).value = "Domicilio"
    xlHoja1.Cells(9, 11).value = "Teléfono"
    xlHoja1.Cells(9, 12).value = "UBICACION"
    xlHoja1.Cells(9, 13).value = "V.CONTRATO"
    xlHoja1.Cells(9, 14).value = "NIVEL"
    xlHoja1.Cells(9, 15).value = "SEXO"
    xlHoja1.Cells(9, 16).value = "AFP"
    xlHoja1.Cells(9, 17).value = "CUSPP"
    
    xlHoja1.Range(xlHoja1.Cells(9, 1), xlHoja1.Cells(9, Col + 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(9, 1), xlHoja1.Cells(9, Col + 14)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(9, 1), xlHoja1.Cells(9, Col + 14)).VerticalAlignment = xlCenter
    'xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas + 3, 1)).Merge True
    xlHoja1.Range(xlHoja1.Cells(9, 1), xlHoja1.Cells(9, Col + 14)).EntireRow.AutoFit
    xlHoja1.Range(xlHoja1.Cells(9, 1), xlHoja1.Cells(9, Col + 14)).WrapText = True

    xlHoja1.Range("A:A").NumberFormat = "@"
    xlHoja1.Range("A:A").VerticalAlignment = xlCenter
    xlHoja1.Range("B:B").NumberFormat = "@"
    xlHoja1.Range("B:B").VerticalAlignment = xlCenter
    xlHoja1.Range("C:C").NumberFormat = "@"
    xlHoja1.Range("C:C").VerticalAlignment = xlCenter
    
    'xlHoja1.Range("F10:IV60000").NumberFormat = "#,##,0.00"
    xlHoja1.Range("F10:AZ60000").NumberFormat = "#,##,0.00"
    
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.Zoom = 65
    
    OLE1.Class = "ExcelWorkSheet"
    ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
    OLE1.SourceDoc = lsArchivoN
    OLE1.Verb = 1
    OLE1.Action = 1
    OLE1.DoVerb -1
End Sub

Private Sub cmdGrabar_Click()
    Dim lsHoy As String
    Dim lsCodPla As String
    Dim lnMontoD As Currency
    Dim lnMontoI As Currency
    Dim lnMontoTI As Currency
    Dim oPla As NActualizaDatosConPlanilla
    Set oPla = New NActualizaDatosConPlanilla
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    
    On Error GoTo ERROR
    
    If Not Valida Then Exit Sub
    If Not lbEditado Then
        lsHoy = Format(gdFecSis, "yyyymmdd")
    Else
        lsHoy = Left(Me.txtPlanillaIns.Text, 8)
    End If
    
    Set rsP = FlexARecordSet(Me.Flex)
    oPla.ModificaPlanilla lsHoy, Me.TxtPlanillas.Text, rsP, GetMovNro(gsCodUser, gsCodAge), Format(CDate(Me.mskFecImp.Text), gsFormatoFecha), Format(CDate(txtFecIni.Text), gsFormatoFecha), Format(CDate(txtFecFin.Text), gsFormatoFecha), Me.txtTpoCambio.Text, Me.txtDes.Text, Format(CDate(Me.mskFecImp.Text), gsFormatoFecha), Me.txtCambFij.Text, Right(Me.cmbEstadoPla.Text, 2)
    
    ClearScreen False
    Activa False, False
    TxtPlanillas.Enabled = True
    ValidaBotones False
    CargaPlanillasTpo TxtPlanillas.Text
    
    txtPlanillaIns.Text = lsHoy & Me.TxtPlanillas.Text
    txtPlanillaIns_EmiteDatos
    
    Exit Sub
ERROR:
    TxtPlanillas.Enabled = True
    ValidaBotones False
    Activa False, False
    MsgBox "Error en la Grabaciòn de la Planilla. No se ha grabado verifique los datos de la planilla.", vbInformation, "Aviso"
End Sub

Private Sub CmdImprimir_Click()
    Dim oPrevio As Previo.clsPrevio
    Dim i As Integer
    Set oPrevio = New Previo.clsPrevio
    Set oPlaEvento = New NActualizaDatosConPlanilla
    Dim oCon As NConstSistemas
    Set oCon = New NConstSistemas
    Dim PB(9) As Boolean 'EJVG 20110818
    Dim lsCadena As String
    Dim lsCadenaExt As String
    
    Dim lsModeloPlantilla As String
       
    If Not IsDate(mskFecImp.Text) Then
        MsgBox "Fecha de Impresión no valida.", vbInformation, "Aviso"
        mskFecImp.SetFocus
        Exit Sub
    End If
    
    If TxtPlanillas.Text = "" Then
        MsgBox "Debe Elegir un tipo de planilla.", vbInformation, "Aviso"
        TxtPlanillas.SetFocus
        Exit Sub
    End If
    
    If txtPlanillaIns.Text = "" Then
        MsgBox "Debe Elegir una planilla.", vbInformation, "Aviso"
        txtPlanillaIns.SetFocus
        Exit Sub
    End If
    
    frmImpreRRHH.Ini "Boletas de Pago;Planilla RRHH;Planilla Contabilidad;Estables/Contratados;Resumen por Agencia;Cartas de Utilidad;Cartas de CTS;Lista de Pagos;Liquidación de Beneficios Sociales;", Caption, PB, gdFecSis, gdFecSis, False 'EJVG 20110818
    
    For i = 1 To Me.Flex.Rows - 1
        If Flex.TextMatrix(i, 0) <> "1" Then
            lsCadenaExt = lsCadenaExt & "'" & Flex.TextMatrix(i, 2) & "',"
        End If
    Next i
    If lsCadenaExt = "" Then
        lsCadenaExt = "''"
    Else
        lsCadenaExt = lsCadenaExt & "''"
    End If
    
    lsCadena = ""
    If PB(1) Then
        lsCadena = lsCadena & oPlaEvento.GetBoletas(Left(Me.txtPlanillaIns, 8), Me.TxtPlanillas.Text, lsCadenaExt, Me.lblPlanillaRes.Caption, Me.mskFecImp.Text, gsRUC, gsEmpresaCompleto, Me.txtDes.Text)
    End If
        
    If PB(2) Then
        'False, ,
        lsCadena = lsCadena & oPlaEvento.GetPlanillas(Left(Me.txtPlanillaIns, 8), Me.TxtPlanillas.Text, lsCadenaExt, Me.lblPlanillaRes.Caption, Me.mskFecImp.Text, gsRUC, gsEmpresaCompleto, Me.txtFecIni.Text, Me.txtFecFin.Text, False, , Me.txtDes.Text)
    End If
    
    If PB(3) Then
        lsCadena = lsCadena & oPlaEvento.GetPlanillas(Left(Me.txtPlanillaIns, 8), Me.TxtPlanillas.Text, lsCadenaExt, Me.lblPlanillaRes.Caption, Me.mskFecImp.Text, gsRUC, gsEmpresaCompleto, Me.txtFecIni.Text, Me.txtFecFin.Text, True, , Me.txtDes.Text)
    End If
    
    If PB(4) Then
        lsCadena = lsCadena & oPlaEvento.GetPlanillas(Left(Me.txtPlanillaIns, 8), Me.TxtPlanillas.Text, lsCadenaExt, Me.lblPlanillaRes.Caption, Me.mskFecImp.Text, gsRUC, gsEmpresaCompleto, Me.txtFecIni.Text, Me.txtFecFin.Text, True, True, Me.txtDes.Text)
    End If
    
    If PB(5) Then
        lsCadena = lsCadena & GetResumenporAgencias(Left(Me.txtPlanillaIns.Text, 8), Me.TxtPlanillas.Text, Me.mskFecImp.Text, gsRUC, gsEmpresaCompleto, Me.txtFecIni.Text, Me.txtFecFin.Text, Me.lblPlanillaRes.Caption, Me.txtDes.Text)
    End If
    
    If PB(6) Then
        lsCadena = lsCadena & oPlaEvento.GetCartas(FlexARecordSet(Me.Flex), gsRUC, gsEmpresaCompleto, Me.txtFecIni.Text, Me.txtFecFin, gdFecSis, Me.mskFecImp.Text)
    End If
    
    If PB(7) Then
        lsCadena = lsCadena & oPlaEvento.GetCartasCTS(FlexARecordSet(Me.Flex), gsRUC, gsEmpresaCompleto, Me.txtFecIni.Text, Me.txtFecFin, Me.mskFecImp.Text, CCur(Me.txtTpoCambio.Text))
    End If
    
    If PB(8) Then
        lsCadena = lsCadena & oPlaEvento.GetListaPagos(TxtPlanillas.Text, Left(Me.txtPlanillaIns.Text, 8), Me.mskFecImp.Text, Me.txtTpoCambio.Text, gsNomAge, gsEmpresa, gdFecSis)
    End If
    
    'EJVG 20110818**************************
    If PB(9) Then
        If Me.TxtPlanillas.Text = "E08" Then
            generarReporteLiquidacionBeneficios
        Else
                MsgBox "Reporte de Liquidación de Planilla no corresponde con la Planilla Actual", vbInformation, "Aviso"
        End If
    End If
    '****************************************
    
    If lsCadena <> "" Then
        If Not PB(1) And Not PB(6) Then
            oPrevio.Show lsCadena, Caption, True, 66, gImpresora
        Else
        '-------------------------------------------------
        If PB(1) Then
                
                Dim MSWord As Word.Application
                Dim MSWordSource As Word.Application
                Set MSWord = New Word.Application
                Set MSWordSource = New Word.Application
                Dim RangeSource As Word.Range
                
                MSWordSource.Documents.Open FileName:=App.path & "\SPOOLER\Boletas_Pago.doc"
                Set RangeSource = MSWordSource.ActiveDocument.Content
                'Lo carga en Memoria
                MSWordSource.ActiveDocument.Content.Copy
                'MSWordSource.ActiveDocument
                'Crea Nuevo Documento
                MSWord.Documents.Add
                
                MSWord.Application.Selection.TypeParagraph
                MSWord.Application.Selection.Paste
                MSWord.Application.Selection.InsertBreak
                
                'MSWordSource.ActiveDocument.Close
                MSWordSource.ActiveDocument.Close
                Set MSWordSource = Nothing
                    
                MSWord.Selection.SetRange start:=MSWord.Selection.start, End:=MSWord.ActiveDocument.Content.End
                MSWord.Selection.MoveEnd
            
                
                MSWord.ActiveDocument.Range.InsertBefore lsCadena
                MSWord.ActiveDocument.Select
                MSWord.ActiveDocument.Range.Font.Name = "Courier New"
                MSWord.ActiveDocument.Range.Font.Size = 6
                MSWord.ActiveDocument.Range.Paragraphs.Space1
                
                MSWord.Selection.Find.Execute Replace:=wdReplaceAll
                MSWord.ActiveDocument.PageSetup.Orientation = wdOrientLandscape
                
                MSWord.ActiveDocument.PageSetup.TopMargin = CentimetersToPoints(2)
                MSWord.ActiveDocument.PageSetup.LeftMargin = CentimetersToPoints(1)
                MSWord.ActiveDocument.PageSetup.RightMargin = CentimetersToPoints(1)

                'Documento.PageSetup.RightMargin = CentimetersToPoints(0.5)
                
                MSWord.ActiveDocument.SaveAs App.path & "\SPOOLER\Boletas_Pago_" & gsCodUser & Format(Now, "yyyymmsshhmmss") & ".doc"
                MSWord.Visible = True
                Set MSWord = Nothing
                Exit Sub
            End If
         '--------------------------------------------------------------------
            If IIf(oCon.LeeConstSistema(gConstSistBoletaRRHH) = "1", True, False) Then
                oPrevio.Show lsCadena, Caption, False, 66, gImpresora
            Else
                oPrevio.Show lsCadena, Caption, True, 33, gImpresora
            End If
        End If
    End If
    gsEmpresaCompleto = Trim(gsEmpresaCompleto)
End Sub

Private Sub generarReporteLiquidacionBeneficios() 'EJVG 20110818
    Dim i, J, k As Integer
    Dim oPla As DInterprete
    Dim oSisPension As DActualizaDatosRRHH
    Dim I_SUE_BAS, I_BON_CAR_FAM, I_RIECAJA, I_PRO_BONOS, I_MOVILI, SEXTA_GRATI As String
    Dim MES_CTS, DIAS_CTS, CTS1, CTS2 As String
    Dim lsCodPers As String
    Dim nemosExcluidos(5) As String
    Dim nemo As String
    Dim esNemoExcluido As Boolean
    Dim Plantilla As String
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    Dim fs As New Scripting.FileSystemObject
    
    'Excluyo estos conceptos tipos descuentos porque se mostraran como retenciones
    nemosExcluidos(1) = "D_SNP_LIQUID"
    nemosExcluidos(2) = "D_D_AFP_P_SEG_LIQ"
    nemosExcluidos(3) = "D_AFP_ASEG_LIQ"
    nemosExcluidos(4) = "D_AFP_C_VAR_LIQ"
    nemosExcluidos(5) = "D_TOT_DESC"
            
    Plantilla = App.path & "\FormatoCarta\FormatoLiquidacionBeneficios.doc"
    If Not fs.FileExists(Plantilla) Then
        MsgBox "No existe el Formato de Liquidación de Beneficios, comuniquese con Sistemas", vbInformation
        Exit Sub
    End If
            
    Set oPla = New DInterprete
    Set oSisPension = New DActualizaDatosRRHH
    
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application
        
    wAppSource.Documents.Open FileName:=Plantilla
    wAppSource.ActiveDocument.Content.Copy
    wApp.Documents.Add
                 
    For i = 1 To Flex.row
        lsCodPers = Flex.TextMatrix(i, 2)
        If lsCodPers <> "" Then
            Dim oPersonaRH As DRHReportes
            Dim rsRH As ADODB.Recordset
            Dim lsNombre As String, lsCargo As String, lsFecIni As String, lsFecFin As String, lsTiempoServicio As String, lsMotivo As String
            Dim totalMen As Double, totalCom As Double
                        
            Set rsRH = New ADODB.Recordset
            Set oPersonaRH = New DRHReportes
            Set rsRH = oPersonaRH.GetDatosPersonaRescindida(lsCodPers)
                        
            If Not (rsRH.BOF Or rsRH.EOF) Then
                lsNombre = rsRH!cPersNombre
                lsCargo = rsRH!cRHCargoDescripcion
                lsFecIni = rsRH!dIngreso
                lsFecFin = rsRH!dCese
                lsTiempoServicio = ""
                lsMotivo = rsRH!cMotivo
                
                I_SUE_BAS = Format(oPla.GetObtenerValorConcepto(lsCodPers, "E01", "I_SUE_BAS"), "#,##0.00")
                I_BON_CAR_FAM = Format(oPla.GetObtenerValorConcepto(lsCodPers, "E01", "I_BON_CAR_FAM"), "#,##0.00")
                I_RIECAJA = Format(oPla.GetObtenerValorConcepto(lsCodPers, "E01", "I_RIECAJA"), "#,##0.00")
                I_PRO_BONOS = Format(oPla.GetObtenerValorConcepto(lsCodPers, "E01", "I_PROM_BON_LIQ"), "#,##0.00")
                I_MOVILI = Format(oPla.GetObtenerValorConcepto(lsCodPers, "E01", "I_MOVILI"), "#,##0.00")
                SEXTA_GRATI = Format(oPla.GetUltimaGratificacion(lsCodPers) / 6, "#,##0.00")
                
                totalMen = CDbl(I_SUE_BAS) + CDbl(I_BON_CAR_FAM) + CDbl(I_RIECAJA) + CDbl(I_MOVILI)
                totalCom = CDbl(I_SUE_BAS) + CDbl(I_BON_CAR_FAM) + CDbl(I_RIECAJA) + CDbl(I_PRO_BONOS) + CDbl(SEXTA_GRATI)
                
                wApp.Application.Selection.TypeParagraph
                wApp.Application.Selection.Paste
                wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
                wApp.Selection.MoveEnd
                
                With wApp.Selection.Find
                        .Text = "<<Nombres>>"
                        .Replacement.Text = PstaNombre(lsNombre, False)
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<FechaIngreso>>"
                        .Replacement.Text = Day(CDate(lsFecIni)) & " DE " & UCase(Format(lsFecIni, "MMMM")) & " DEL " & Year(CDate(lsFecIni))
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<FechaCese>>"
                        .Replacement.Text = Day(CDate(lsFecFin)) & " DE " & UCase(Format(lsFecFin, "MMMM")) & " DEL " & Year(CDate(lsFecFin))
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                lsTiempoServicio = obtieneDifFechasEnAñoMesDia(CDate(lsFecIni), CDate(lsFecFin))
                
                With wApp.Selection.Find
                        .Text = "<<TiempoServicio>>"
                        .Replacement.Text = lsTiempoServicio
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<Motivo>>"
                        .Replacement.Text = lsMotivo
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<Cargo>>"
                        .Replacement.Text = lsCargo
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                With wApp.Selection.Find
                        .Text = "<<SuelMen>>"
                        .Replacement.Text = I_SUE_BAS
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<SuelCom>>"
                        .Replacement.Text = I_SUE_BAS
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<AsigMen>>"
                        .Replacement.Text = I_BON_CAR_FAM
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<AsigCom>>"
                        .Replacement.Text = I_BON_CAR_FAM
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<RCMen>>"
                        .Replacement.Text = "0.00"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<RCCom>>"
                        .Replacement.Text = I_RIECAJA
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<PBMen>>"
                        .Replacement.Text = "0.00"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<PBCom>>"
                        .Replacement.Text = I_PRO_BONOS
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<MovMens>>"
                        .Replacement.Text = I_MOVILI
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<MovCom>>"
                        .Replacement.Text = "0.00"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<SGMen>>"
                        .Replacement.Text = "0.00"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<SGCom>>"
                        .Replacement.Text = SEXTA_GRATI
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<TotMen>>"
                        .Replacement.Text = Format(totalMen, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<TotCom>>"
                        .Replacement.Text = Format(totalCom, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                Dim diasCTS, lnMesSistema As Integer
                Dim I_BON_PROD, I_LIQCTS, lsPeridoIniCTS As String
                diasCTS = oPla.GetObtenerValorConcepto(lsCodPers, "E08", "U_D_T_CTS")
                MES_CTS = diasCTS \ 30
                DIAS_CTS = diasCTS Mod 30
                CTS1 = totalCom / 12 * MES_CTS
                CTS2 = totalCom / 12 / 30 * DIAS_CTS
                
                lnMesSistema = Month(gdFecSis)
                If lnMesSistema >= 5 And lnMesSistema < 11 Then
                    lsPeridoIniCTS = "01/05/" & Year(gdFecSis)
                Else
                    If lnMesSistema = 11 Or lnMesSistema = 12 Then
                        lsPeridoIniCTS = "01/11/" & Year(gdFecSis)
                    Else
                        lsPeridoIniCTS = "01/11/" & Year(DateAdd("YYYY", -1, gdFecSis))
                    End If
                End If
                
                With wApp.Selection.Find
                        .Text = "<<FechaCTS>>"
                        .Replacement.Text = lsPeridoIniCTS
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<FechaCeseN>>"
                        .Replacement.Text = lsFecFin
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<MesCTS>>"
                        .Replacement.Text = MES_CTS
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<DiasCTS>>"
                        .Replacement.Text = DIAS_CTS
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<CTS1>>"
                        .Replacement.Text = Format(CTS1, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<CTS2>>"
                        .Replacement.Text = Format(CTS2, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                Dim lsNemoCabecera As String
                For k = 0 To Flex.Cols - 1
                    lsNemoCabecera = Flex.TextMatrix(0, k)
                    If lsNemoCabecera = "I_BON_PROD_LIQUI" Then
                        I_BON_PROD = Format(Flex.TextMatrix(i, k), "#,##0.00")
                    ElseIf lsNemoCabecera = "I_LIQCTS" Then
                        I_LIQCTS = Format(Flex.TextMatrix(i, k), "#,##0.00")
                    End If
                Next
                
                With wApp.Selection.Find
                        .Text = "<<BON_PROD>>"
                        .Replacement.Text = I_BON_PROD
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<I_LIQCTS>>"
                        .Replacement.Text = I_LIQCTS
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                Dim diasREMU As Integer
                Dim TotMenSinPromedioBono, I_REM_D_CTS As String
                
                diasREMU = oPla.GetObtenerValorConcepto(lsCodPers, "E08", "U_D_T_M")
                TotMenSinPromedioBono = totalMen - CDbl(I_MOVILI)
                
                With wApp.Selection.Find
                        .Text = "<<DiasTM>>"
                        .Replacement.Text = diasREMU
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<TM-Mov>>"
                        .Replacement.Text = Format(TotMenSinPromedioBono, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<2A1>>"
                        .Replacement.Text = Format(TotMenSinPromedioBono / 30 * diasREMU, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<2A2>>"
                        .Replacement.Text = Format(I_MOVILI / 30 * diasREMU, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                For k = 0 To Flex.Cols - 1
                    If Flex.TextMatrix(0, k) = "I_REM_D_CTS" Then
                        I_REM_D_CTS = Format(Flex.TextMatrix(i, k), "#,##0.00")
                        Exit For
                    End If
                Next
                
                With wApp.Selection.Find
                        .Text = "<<I_REM_D_CTS>>"
                        .Replacement.Text = I_REM_D_CTS
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                'VACACIONES
                '*****************************
                Dim diasVACA As Integer
                Dim I_VAC_LIQUID, TotComSinSextaGrat As String
                diasVACA = oPla.GetObtenerValorConcepto(lsCodPers, "E08", "U_DIAS_VACA")
                TotComSinSextaGrat = totalCom - CDbl(SEXTA_GRATI)
                
                With wApp.Selection.Find
                        .Text = "<<TotCom-SGCom>>"
                        .Replacement.Text = Format(TotComSinSextaGrat, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<DiasVaca>>"
                        .Replacement.Text = diasVACA
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<2B1>>"
                        .Replacement.Text = Format(TotComSinSextaGrat / 30 * diasVACA, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                For k = 0 To Flex.Cols - 1
                    If Flex.TextMatrix(0, k) = "I_VAC_LIQUID" Then
                        I_VAC_LIQUID = Format(Flex.TextMatrix(i, k), "#,##0.00")
                        Exit For
                    End If
                Next
                           
                With wApp.Selection.Find
                        .Text = "<<I_VAC_LIQUID>>"
                        .Replacement.Text = I_VAC_LIQUID
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                           
                'VACACIONES TRUNCAS
                '*****************************
                Dim diasVACATRUN As Integer
                Dim I_LIQVACTRUN, MES_VAC_TRUN, DIAS_VAC_TRUN As String
                
                I_LIQVACTRUN = ""
                MES_VAC_TRUN = ""
                DIAS_VAC_TRUN = ""
                diasVACATRUN = oPla.GetObtenerValorConcepto(lsCodPers, "E08", "U_D_VACA_TRUN")
                MES_VAC_TRUN = diasVACATRUN \ 30
                DIAS_VAC_TRUN = diasVACATRUN Mod 30
                
                
                With wApp.Selection.Find
                        .Text = "<<PerTrunca>>"
                        .Replacement.Text = MES_VAC_TRUN & " meses," & DIAS_VAC_TRUN & " días"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<MesTrun>>"
                        .Replacement.Text = MES_VAC_TRUN
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<DiaTrun>>"
                        .Replacement.Text = DIAS_VAC_TRUN
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<2C1>>"
                        .Replacement.Text = Format(TotComSinSextaGrat / 12 * MES_VAC_TRUN, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<2C2>>"
                        .Replacement.Text = Format(TotComSinSextaGrat / 12 / 30 * DIAS_VAC_TRUN, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                For k = 0 To Flex.Cols - 1
                    If Flex.TextMatrix(0, k) = "I_LIQVACTRUN" Then
                        I_LIQVACTRUN = Format(Flex.TextMatrix(i, k), "#,##0.00")
                        Exit For
                    End If
                Next
                
                With wApp.Selection.Find
                        .Text = "<<I_LIQVACTRUN>>"
                        .Replacement.Text = I_LIQVACTRUN
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                'GRATIFICACION TRUNCA
                '******************************
                Dim I_LIQGRATIF, MES_GRAT_TRUN As String
                MES_GRAT_TRUN = oPla.GetObtenerValorConcepto(lsCodPers, "E08", "U_M_GRATI_TRUN")
                
                With wApp.Selection.Find
                        .Text = "<<MesGT>>"
                        .Replacement.Text = MES_GRAT_TRUN
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                        .Text = "<<2D1>>"
                        .Replacement.Text = Format(TotComSinSextaGrat / 6 * MES_GRAT_TRUN, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                For k = 0 To Flex.Cols - 1
                    If Flex.TextMatrix(0, k) = "I_LIQGRATIF" Then
                        I_LIQGRATIF = Format(Flex.TextMatrix(i, k), "#,##0.00")
                        Exit For
                    End If
                Next
                
                With wApp.Selection.Find
                        .Text = "<<I_LIQGRATIF>>"
                        .Replacement.Text = I_LIQGRATIF
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                'BONIFICACION EXTRAORDINARIA
                '***********************************
                Dim EPS_ESSALUD, BONIFICACION, CANASTA_NAVIDEÑA, DEV_QUINTA_CAT As String
                BONIFICACION = IIf(oPla.GetCodAsistenciaMedica(lsCodPers) = 25, "I_B_E_ES_CTS", "I_B_E_EPS_CTS")
                
                For k = 0 To Flex.Cols - 1
                    If Flex.TextMatrix(0, k) = BONIFICACION Then
                        EPS_ESSALUD = Format(Flex.TextMatrix(i, k), "#,##0.00")
                        Exit For
                    End If
                Next
                
                With wApp.Selection.Find
                        .Text = "<<EPS_ESSALUD>>"
                        .Replacement.Text = EPS_ESSALUD
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                For k = 0 To Flex.Cols - 1
                    If Flex.TextMatrix(0, k) = "I_CANAST_NAVIDEÑA" Then
                        CANASTA_NAVIDEÑA = Format(Flex.TextMatrix(i, k), "#,##0.00")
                        Exit For
                    End If
                Next
                
                With wApp.Selection.Find
                        .Text = "<<CANASTA_NAV>>"
                        .Replacement.Text = CANASTA_NAVIDEÑA
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                For k = 0 To Flex.Cols - 1
                    If Flex.TextMatrix(0, k) = "I_DEVQUINTACAT" Then
                        DEV_QUINTA_CAT = Format(Flex.TextMatrix(i, k), "#,##0.00")
                        Exit For
                    End If
                Next
                
                With wApp.Selection.Find
                        .Text = "<<DEV_5TA_CAT>>"
                        .Replacement.Text = DEV_QUINTA_CAT
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                'TOTAL DE BENEFICIOS Y TOTAL DE INGRESOS
                '***************************************************
                Dim I_TOT_ING As String
                With wApp.Selection.Find
                        .Text = "<<BENEFICIOS>>"
                        .Replacement.Text = Format(CDbl(IIf(I_REM_D_CTS = "", 0, I_REM_D_CTS)) + CDbl(IIf(I_BON_PROD = "", 0, I_BON_PROD)) + CDbl(IIf(I_VAC_LIQUID = "", 0, I_VAC_LIQUID)) + CDbl(IIf(I_LIQVACTRUN = "", 0, I_LIQVACTRUN) + CDbl(IIf(I_LIQGRATIF = "", 0, I_LIQGRATIF)) + CDbl(IIf(EPS_ESSALUD = "", 0, EPS_ESSALUD)) + CDbl(IIf(CANASTA_NAVIDEÑA = "", 0, CANASTA_NAVIDEÑA)) + CDbl(IIf(DEV_QUINTA_CAT = "", 0, DEV_QUINTA_CAT))), "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                For k = 0 To Flex.Cols - 1
                    If Flex.TextMatrix(0, k) = "I_TOT_ING" Then
                        I_TOT_ING = Format(Flex.TextMatrix(i, k), "#,##0.00")
                        Exit For
                    End If
                Next
                
                With wApp.Selection.Find
                        .Text = "<<I_TOT_ING>>"
                        .Replacement.Text = I_TOT_ING
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                'DESCUENTOS
                '*************************************
                Dim oConRH As DRHConcepto
                Dim cont, m As Integer
                Dim valorNemo, totalNemoDescuento As Double
                
                Set oConRH = New DRHConcepto
                esNemoExcluido = False
                cont = 0
                totalNemoDescuento = 0
                
                For k = 5 To Flex.Cols - 1
                    nemo = Flex.TextMatrix(0, k)
                    If Left(nemo, 2) = "D_" Then
                        For m = 1 To UBound(nemosExcluidos)
                            If nemo = nemosExcluidos(m) Then
                                esNemoExcluido = True
                                Exit For
                            End If
                        Next
                        If esNemoExcluido = False Then
                            valorNemo = Flex.TextMatrix(i, k)
                            If valorNemo <> 0 Then
                                cont = cont + 1
                                totalNemoDescuento = totalNemoDescuento + valorNemo
                                With wApp.Selection.Find
                                    .Text = "<<DESC" & cont & ">>"
                                    .Replacement.Text = oConRH.GetImpreConcepto(nemo)
                                    .Forward = True
                                    .Wrap = wdFindContinue
                                    .Format = False
                                    .Execute Replace:=wdReplaceAll
                                End With
                                With wApp.Selection.Find
                                    .Text = "<<3A" & cont & ">>"
                                    .Replacement.Text = Format(valorNemo, "#,##0.00")
                                    .Forward = True
                                    .Wrap = wdFindContinue
                                    .Format = False
                                    .Execute Replace:=wdReplaceAll
                                End With
                            End If
                        End If
                        esNemoExcluido = False
                    End If
                Next
                'Limpiamos los campos de descuento no utilizados
                Dim n As Integer
                For n = cont + 1 To 10
                    With wApp.Selection.Find
                        .Text = "- <<DESC" & n & ">>"
                        .Replacement.Text = ""
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    With wApp.Selection.Find
                        .Text = "<<3A" & n & ">>"
                        .Replacement.Text = ""
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                    End With
                Next
                'TOTAL DE DESCUENTOS
                '********************************************
                With wApp.Selection.Find
                        .Text = "<<DESC_TOTAL>>"
                        .Replacement.Text = Format(totalNemoDescuento, "#,##0.00")
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                'RETENCIONES
                '********************************************
                Dim rsSP As Recordset
                Dim lnSistemaPension As Integer
                Dim lsAFP_SNP As String
                Set rsSP = New Recordset
                
                Set rsSP = oSisPension.GetRRHHFilePensiones(lsCodPers)
                If Not (rsSP.BOF Or rsSP.EOF) Then
                    lnSistemaPension = rsSP!nConsValor
                    If lnSistemaPension = 0 Then 'AFP
                        lsAFP_SNP = Left(rsSP(3), InStr(1, rsSP(3), "-") - 1)
                    ElseIf lnSistemaPension = 1 Then 'SNP
                        lsAFP_SNP = rsSP!cConsDescripcion
                    End If
                End If
                
                With wApp.Selection.Find
                        .Text = "<<SisPens>>"
                        .Replacement.Text = lsAFP_SNP
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
                
                Dim BaseAFP_SNP, AFP_NSP As String
                Dim lnTotalRetencion, PorcentajeValor As Double
                
                BaseAFP_SNP = I_TOT_ING - (CDbl(IIf(I_LIQCTS = "", 0, I_LIQCTS)) + CDbl(IIf(I_LIQGRATIF = "", 0, I_LIQGRATIF)) + CDbl(IIf(EPS_ESSALUD = "", 0, EPS_ESSALUD)) + (CDbl(IIf(I_MOVILI = "", 0, I_MOVILI)) / 30 * diasREMU))
                lnTotalRetencion = 0
                If lnSistemaPension = 0 Then
                    For k = 1 To 3
                        With wApp.Selection.Find
                            .Text = "<<RET" & k & ">>"
                            If k = 1 Then
                                .Replacement.Text = "- Comisión Asegurable"
                            ElseIf k = 2 Then
                                .Replacement.Text = "- Prima de Seguro"
                            ElseIf k = 3 Then
                                .Replacement.Text = "- Comisión Variable"
                            End If
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                        End With
                        
                        If k = 1 Then
                            PorcentajeValor = CDbl(oPla.FunGetConstante("V_AFP_CUO_FIJ"))
                            For m = 5 To Flex.Cols - 1
                                If Flex.TextMatrix(0, m) = "D_AFP_ASEG_LIQ" Then
                                    AFP_NSP = Flex.TextMatrix(i, m)
                                    Exit For
                                End If
                            Next
                        ElseIf k = 2 Then
                            PorcentajeValor = CDbl(oPla.FunGetValorConstante("GET_AFP_PRIMA", lsCodPers, CDate(Me.txtFecIni.Text), CDate(Me.txtFecFin.Text), "", ""))
                            For m = 5 To Flex.Cols - 1
                                If Flex.TextMatrix(0, m) = "D_D_AFP_P_SEG_LIQ" Then
                                    AFP_NSP = Flex.TextMatrix(i, m)
                                    Exit For
                                End If
                            Next
                        ElseIf k = 3 Then
                            PorcentajeValor = CDbl(oPla.FunGetValorConstante("GET_AFP_COM_VAR", lsCodPers, CDate(Me.txtFecIni.Text), CDate(Me.txtFecFin.Text), "", ""))
                            For m = 5 To Flex.Cols - 1
                                If Flex.TextMatrix(0, m) = "D_AFP_C_VAR_LIQ" Then
                                    AFP_NSP = Flex.TextMatrix(i, m)
                                    Exit For
                                End If
                            Next
                        End If
                                                
                        With wApp.Selection.Find
                            .Text = "<<PORC" & k & ">>"
                            .Replacement.Text = Format(CStr(PorcentajeValor * 100), "#,##0.00") & "%"
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                        End With
                        With wApp.Selection.Find
                            .Text = "<<Sob" & k & ">>"
                            .Replacement.Text = "Sobre"
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                        End With
                        With wApp.Selection.Find
                            .Text = "<<BSP" & k & ">>"
                            .Replacement.Text = Format(BaseAFP_SNP, "#,##0.00")
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                        End With
                        With wApp.Selection.Find
                            .Text = "<<3B" & k & ">>"
                            .Replacement.Text = Format(AFP_NSP, "#,##0.00")
                            lnTotalRetencion = lnTotalRetencion + CDbl(IIf(AFP_NSP = "", 0, AFP_NSP))
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                        End With
                    Next
                ElseIf lnSistemaPension = 1 Then
                    For k = 1 To 3
                        With wApp.Selection.Find
                            .Text = "<<RET" & k & ">>"
                            If k = 1 Then
                                .Replacement.Text = "SNP"
                            ElseIf k = 2 Then
                                .Replacement.Text = ""
                            ElseIf k = 3 Then
                                .Replacement.Text = ""
                            End If
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                        End With
                        
                        If k = 1 Then
                            PorcentajeValor = CDbl(oPla.FunGetConstante("V_POR_DSCT_0NP"))
                            For m = 5 To Flex.Cols - 1
                                If Flex.TextMatrix(0, m) = "D_SNP_LIQUID" Then
                                    AFP_NSP = Flex.TextMatrix(i, m)
                                    Exit For
                                End If
                            Next
                        End If
                        
                        With wApp.Selection.Find
                            .Text = "<<PORC" & k & ">>"
                            If k = 1 Then
                                PorcentajeValor = Format(CStr(PorcentajeValor * 100), "#,##0.00") & "%"
                            ElseIf k = 2 Then
                                .Replacement.Text = ""
                            ElseIf k = 3 Then
                                .Replacement.Text = ""
                            End If
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                        End With
                        With wApp.Selection.Find
                            .Text = "<<Sob" & k & ">>"
                            If k = 1 Then
                                .Replacement.Text = "Sobre"
                            ElseIf k = 2 Then
                                .Replacement.Text = ""
                            ElseIf k = 3 Then
                                .Replacement.Text = ""
                            End If
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                        End With
                        With wApp.Selection.Find
                            .Text = "<<BSP" & k & ">>"
                            If k = 1 Then
                                .Replacement.Text = Format(BaseAFP_SNP, "#,##0.00")
                            ElseIf k = 2 Then
                                .Replacement.Text = ""
                            ElseIf k = 3 Then
                                .Replacement.Text = ""
                            End If
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                        End With
                        With wApp.Selection.Find
                            .Text = "<<3B" & k & ">>"
                            If k = 1 Then
                                .Replacement.Text = AFP_NSP
                                lnTotalRetencion = lnTotalRetencion + CDbl(IIf(AFP_NSP = "", 0, AFP_NSP))
                            ElseIf k = 2 Then
                                .Replacement.Text = ""
                            ElseIf k = 3 Then
                                .Replacement.Text = ""
                            End If
                            
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                        End With
                    Next
                End If
                With wApp.Selection.Find
                    .Text = "<<RET_TOTAL>>"
                    .Replacement.Text = Format(CStr(lnTotalRetencion), "#,##0.00")
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                'TOTAL DESCUENTOS
                '****************************************************
                Dim D_TOT_DESC As String
                Dim ln3A3B, ln123 As Double
                ln3A3B = lnTotalRetencion + totalNemoDescuento
                
                For k = 0 To Flex.Cols - 1
                    If Flex.TextMatrix(0, k) = "D_TOT_DESC" Then
                        D_TOT_DESC = Format(Flex.TextMatrix(i, k), "#,##0.00")
                        Exit For
                    End If
                Next
                
                ln123 = CDbl(IIf(I_TOT_ING = "", 0, I_TOT_ING)) - CDbl(IIf(D_TOT_DESC = "", 0, D_TOT_DESC))
                
                With wApp.Selection.Find
                    .Text = "<<D_TOT_DES>>"
                    .Replacement.Text = D_TOT_DESC
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                    .Text = "<<1+2-3>>"
                    .Replacement.Text = Format(CStr(ln123), "#,##0.00")
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                Dim APOR_ESSALUD, APOR_EPS As String
                
                For k = 0 To Flex.Cols - 1
                    If Flex.TextMatrix(0, k) = "A_SEG_SOCIAL 9" Then
                        APOR_ESSALUD = Format(Flex.TextMatrix(i, k), "#,##0.00")
                    ElseIf Flex.TextMatrix(0, k) = "A_APORT_EPS" Then
                        APOR_EPS = Format(Flex.TextMatrix(i, k), "#,##0.00")
                    End If
                Next
                
                With wApp.Selection.Find
                    .Text = "<<APOREPS>>"
                    .Replacement.Text = APOR_EPS
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                With wApp.Selection.Find
                    .Text = "<<APORESSALUD>>"
                    .Replacement.Text = APOR_ESSALUD
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                With wApp.Selection.Find
                    .Text = "<<SUMEPSESSA>>"
                    .Replacement.Text = Format(CDbl(IIf(APOR_EPS = "", 0, APOR_EPS)) + CDbl(IIf(APOR_ESSALUD = "", 0, APOR_ESSALUD)), "#,##0.00")
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                'DECLARACIÓN RECIBIDO
                '*******************************************
                gsSimbolo = gcMN
                With wApp.Selection.Find
                    .Text = "<<SLIQUI_FINAL>>"
                    .Replacement.Text = ConvNumLet(Format(ln123, "#,##0.00"), False)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With wApp.Selection.Find
                    .Text = "<<FechaSistema>>"
                    .Replacement.Text = "Iquitos, " & Day(gdFecSis) & " de " & Format(gdFecSis, "MMMM") & " del " & Year(gdFecSis)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
            End If
        End If
    Next
    Screen.MousePointer = 0
       
    wAppSource.ActiveDocument.Close
    wApp.ActiveDocument.CopyStylesFromTemplate (Plantilla)
    'wApp.ActiveDocument.SaveAs (App.path & "\Spooler\BeneficioSocial_" & NroCta & "_" & Replace(Left(Time, 5), ":", "") & ".doc")
    wApp.Visible = True
    Set wAppSource = Nothing
    Set wApp = Nothing
End Sub

Private Function obtieneDifFechasEnAñoMesDia(pdFecIni As Date, pdFecFin As Date) As String 'EJVG 20110818
    Dim ldFecIni, ldFecFin As Date
    Dim lnAnio, lnMes, lnDia As Integer
    
    If DateDiff("d", pdFecIni, pdFecFin) >= 0 Then
        ldFecIni = pdFecIni
        ldFecFin = pdFecFin
    Else
        ldFecIni = pdFecFin
        ldFecFin = pdFecIni
    End If
    
    lnAnio = Year(ldFecFin) - Year(ldFecIni)
    lnMes = Month(ldFecFin) - Month(ldFecIni)
    lnDia = Day(ldFecFin) - Day(ldFecIni)
    
    If lnDia < 1 Then
        lnDia = lnDia + 31
        lnMes = lnMes - 1
        Select Case Month(ldFecFin)
            Case 5, 7, 10, 12
                lnDia = lnDia - 1
            Case 3
                lnDia = lnDia - 3
                If (Year(ldFecFin) Mod 4) = 0 Then
                    lnDia = lnDia + 1
                    If (Year(ldFecFin) Mod 100) = 0 And (Year(ldFecFin) Mod 400) <> 0 Then
                        lnDia = lnDia - 1
                    End If
                End If
        End Select
    End If
    If lnMes < 1 Then
        lnMes = lnMes + 12
        lnAnio = lnAnio - 1
    End If
    obtieneDifFechasEnAñoMesDia = FillNum(CStr(lnAnio), 2, "0") & " Años, " & FillNum(CStr(lnMes), 2, "0") & " meses, " & FillNum(CStr(lnDia), 2, "0") & "días"
End Function

Private Sub cmdModificar_Click()
    Dim oPla As DRHProcesosCierre
    Set oPla = New DRHProcesosCierre
    
    If TxtPlanillas.Text = "" Then
        MsgBox "Debe elegir un tipo de Planilla.", vbInformation, "Aviso"
        TxtPlanillas.SetFocus
        Exit Sub
    ElseIf txtPlanillaIns.Text = "" Then
        MsgBox "Debe elegir una Planilla para poder Modificarla.", vbInformation, "Aviso"
        txtPlanillaIns.SetFocus
        Exit Sub
    ElseIf Right(cmbEstadoPla.Text, 1) = RHPlanillaEstado.RHPlanillaEstadoPagado Then
        MsgBox "La Planilla no se puede modificar, porque tiene estado PAGADO", vbInformation, "Aviso"
        'txtPlanillaIns.SetFocus
        Exit Sub
    End If
    
    If oPla.GetContadorPago(Trim(TxtPlanillas.Text), Left(txtPlanillaIns.Text, 8)) > 0 Then
        MsgBox "La Planilla no se puede modificar, porque tiene Trabajadores con Pago", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'MAVM 20110715 ***
    If TxtPlanillas.Text = "E06" Then
        Dim objD As DActualizaDatosConPlanilla
        Set objD = New DActualizaDatosConPlanilla
        Dim i As Integer
        For i = 1 To Me.Flex.Rows - 2
            objD.ActualizarDiasVacaciones Flex.TextMatrix(i, 2), Flex.TextMatrix(i, 5), 1, Left(txtPlanillaIns.Text, 8)
        Next i
    End If
    '***
        
    lbEditado = True
    ValidaBotones True
    
    Activa False, True
    
    cmbEstadoPla.SetFocus
End Sub

Private Sub CmdNuevo_Click()
    Dim oTipo As nTipoCambio
    Set oTipo = New nTipoCambio
     
    If Me.TxtPlanillas.Text = "" Then
        MsgBox "Debe elegir una Planilla para poder generar una Instancia.", vbInformation, "Aviso"
        TxtPlanillas.SetFocus
        Exit Sub
    End If

    IniFlex True

    GetTipCambio gdFecSis, Not gbBitCentral

    Me.txtTpoCambio.Text = Format(gnTipCambioV, "#.###")
    Me.txtCambFij.Text = Format(gnTipCambio, "#.###")

    lbEditado = False
    ClearScreen False
    Activa False, True
    ValidaBotones True

    cmbEstadoPla.Enabled = False
    cmbEstadoPla.ListIndex = 0
    TxtPlanillas.Enabled = False

    If TxtPlanillas.Text = gsRHPlanillaSubsidio Then
        txtFecIni = Format(gdFecSis, gsFormatoFechaView)
        txtFecFin = Format(gdFecSis, gsFormatoFechaView)
        txtFecIni.Enabled = False
        txtFecFin.Enabled = False
    End If

End Sub

Private Sub cmdProcesar_Click()
    Dim i As Long
    Dim J As Integer
    Dim lsCodPers As String
    Dim lsCodEmp As String
    Dim lsConcepto As String
    Dim ldFecIni As Date
    Dim ldFecFin As Date
    Dim lsResultado As String
    Dim lsCodPla As String
    Dim lnPorIngresos As Double
    Dim lnPorDiaLab As Double
    Dim oPla As DInterprete
    Set oPla = New DInterprete
    
    lbCancela = False
    If MsgBox("Desea Procesar la Planilla ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oPla.Interprete_InI
    
    If Not Valida Then Exit Sub
    
    lnSalir = 1
    ValidaProceso False
    
    ldFecIni = CDate(txtFecIni)
    ldFecFin = CDate(txtFecFin)
    lsCodPla = TxtPlanillas.Text
    
    IniTValor True
    Me.cmdCancelarPlanilla.Enabled = False
    oPlaEvento_ShowProgress
    For i = 1 To Me.Flex.Rows - 1
        
        If Flex.TextMatrix(i, 0) = "1" And Left(Flex.TextMatrix(i, 1), 1) <> "_" Then
            oPla.IniTValor False
            
            If lbCancela Then
                MsgBox "El Proceso de calculo ha sido cancelado.", vbInformation, "Aviso"
                lnSalir = 0
                ValidaProceso True
                cmdCancelarPlanilla.Enabled = True
                Exit Sub
            End If
            
            For J = 5 To Me.Flex.Cols - 1
                lsCodEmp = Flex.TextMatrix(i, 1)
                lsCodPers = Flex.TextMatrix(i, 2)
                lsConcepto = Trim(Flex.TextMatrix(0, J))
                
                oPlaEvento_Progress i, Flex.Rows - 1
                
                If Left(lsConcepto, 1) <> "_" Then
                    Caption = "Planilla: " & lsCodEmp & " - " & lsConcepto
                    If oPla.EsConceptoEmpleado(lsConcepto, lsCodPers, lsCodPla) Then
                        'ldFecIni = CDate(txtFecIni)
                        lsResultado = oPla.FunFiltro(lsConcepto)
                        lsResultado = oPla.GetValorFunLog(lsResultado, lsCodPers, ldFecIni, ldFecFin, True, TxtPlanillas.Text, Left(Me.txtPlanillaIns, 8))
                        lsResultado = oPla.FunEvalua(lsResultado, lsCodPers, ldFecIni, ldFecFin, True, TxtPlanillas.Text, Left(Me.txtPlanillaIns, 8))
                        DoEvents
                    Else
                        lsResultado = "0"
                    End If
                    If lsResultado <> "" Then
                        If Left(lsConcepto, 2) <> "U_" Then
                            Flex.TextMatrix(i, J) = Format((ExprANum(lsResultado)), "#,##0.00")
                        Else
                            Flex.TextMatrix(i, J) = Format(CCur((ExprANum(lsResultado))), "#,##0.000")
                        End If
                    Else
                        Flex.TextMatrix(i, J) = Format(0, "#,##0.00")
                    End If
                    
                    If oPla.EsTotalPlanilla(lsConcepto) Then
                        Flex.row = i
                        Flex.Col = J
                        Flex.CellBackColor = &HC0C000
                        Flex.CellFontBold = True
                    End If
                End If
            Next J
        End If
    Next i
    oPlaEvento_CloseProgress
    Me.Caption = "Generación de Planillas"
    lnSalir = 0
    ValidaProceso True
    cmdCancelarPlanilla.Enabled = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdValida_Click()
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    Dim lsCadena As String
    
    Dim rsPla As ADODB.Recordset
    Set rsPla = New ADODB.Recordset
    
    lsCadena = oPla.GetErroresPlanillas(Me.TxtPlanillas.Text, Left(txtPlanillaIns.Text, 8), "", gbBitCentral)

    If lsCadena <> "" Then lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    Set rsPla = FlexARecordSet(Flex)
    While Not rsPla.EOF
        'MAVM 20120214 Fractal***
        'If (rsPla.Fields(gsRHConceptoITOTING) - rsPla.Fields(gsRHConceptoDTOTDES)) < 0 Then
        If rsPla.Fields(gsRHConceptoINETOPAGAR) < 0 Then
            lsCadena = lsCadena & rsPla!nombre & " . Tiene un Neto a Pagar Negativo." & oImpresora.gPrnSaltoLinea
        End If
        rsPla.MoveNext
    Wend
    
    'Set rsPla = FlexARecordSet(Flex)
    'While Not rsPla.EOF
    '    If (prRs.Fields(gsRHConceptoITOTING) - prRs.Fields(gsRHConceptoDTOTDES)) < 0 Then
    '        lsCadena = lsCadena & prRs!Nombre & " . Tiene un Neto a Pagar Negativo." & oImpresora.gPrnSaltoLinea
    '    End If
    '    rsPla.MoveNext
    'Wend
    
    If lsCadena = "" Then
        MsgBox "Planilla sin Observaciones.", vbInformation, "Aviso"
    Else
        oPrevio.Show lsCadena, "Errores de la Planilla"
    End If
End Sub

Private Sub Flex_DblClick()
    If Flex.TextMatrix(Flex.row, 3) = "TOTAL" Then Exit Sub
    If Flex.Col = 4 Then
        If Flex.TextMatrix(Flex.row, 0) = "1" Then
            Flex.TextMatrix(Flex.row, 0) = "0"
            Set Flex.CellPicture = picNo
        Else
            Flex.TextMatrix(Flex.row, 0) = "1"
            Set Flex.CellPicture = picSi
        End If
    End If
End Sub

Private Sub FLEX_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdProcesar.Enabled Then
            Me.cmdProcesar.SetFocus
        Else
                
        End If
    ElseIf KeyAscii = 32 Then
        Flex_DblClick
    End If
End Sub

Private Sub Flex_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnu
End Sub

Private Sub Form_Load()
    Dim oPla As DActualizaDatosConPlanilla
    Dim oCon As DRHConcepto
    Dim rsP As ADODB.Recordset
    Dim oCons As DConstantes
    Set oPla = New DActualizaDatosConPlanilla
    Set oCon = New DRHConcepto
    Set rsP = New ADODB.Recordset
    Set oCons = New DConstantes
    lnSalir = 0
    lbCancela = False
    cmbEstadoPla.Enabled = False
    
    cmdProcesar.Enabled = False
    Interprete_InI
    TxtPlanillas.rs = oPla.GetPlanillas(, True)
     
    Flex.Rows = 1
    Flex.Rows = 2
    Flex.Cols = 5
    Flex.FixedRows = 1
    Flex.FixedCols = 4
    
    Flex.TextMatrix(0, 0) = "B"
    Flex.TextMatrix(0, 1) = "Cod.Emp."
    Flex.TextMatrix(0, 2) = "CodPers"
    Flex.TextMatrix(0, 3) = "Nombre"
    Flex.TextMatrix(0, 4) = "OK"
    
    Flex.ColWidth(0) = 1
    Flex.ColWidth(1) = 900
    Flex.ColWidth(2) = 1
    Flex.ColWidth(3) = 3900
    Flex.ColWidth(4) = 400
    fraPla.Enabled = False
    
    Set rsP = oCons.GetConstante(6039)
    CargaCombo rsP, Me.cmbOpc
    rsP.Close
    
    Set rsP = oCons.GetConstante(6040)
    CargaCombo rsP, Me.cmbEstadoPla
    cmbOpc.ListIndex = 2
    
    mskFecImp.Text = Format(gdFecSis, gsFormatoFechaView)
    
    rsP.Close
    Set rsP = Nothing
    Set oPla = Nothing
    Set oCon = Nothing
    Set oCons = Nothing
    
    If lnTipo = gTipoProcesoRRHHAbono Then
        Me.cmdAbonarCuentas.Enabled = True
        Me.cmdValida.Enabled = True
        Me.cmdAbonarCuentas.Visible = True
        Me.cmdValida.Visible = True
        Me.cmdCancelar.Visible = False
        Me.cmdEliminar.Visible = False
        Me.cmdGrabar.Visible = False
        Me.cmdNuevo.Visible = False
        Me.cmdImprimir.Visible = False
        Me.cmdModificar.Visible = False
        Me.cmdCancelarPlanilla.Visible = False
        Me.cmdProcesar.Visible = False
        cmdexportar.Visible = False
        Me.cmdAsiento.Visible = True '
    ElseIf lnTipo = gTipoProcesoRRHHCalculo Then
        Me.cmdAbonarCuentas.Visible = False
        Me.cmdAsiento.Visible = False
        Me.cmdValida.Visible = False
        cmdexportar.Enabled = True
    ElseIf lnTipo = gTipoProcesoRRHHConsulta Then
        Me.cmdAbonarCuentas.Enabled = True
        Me.cmdValida.Enabled = True
        Me.cmdAbonarCuentas.Visible = True
        Me.cmdValida.Visible = True
        Me.cmdCancelar.Visible = False
        Me.cmdEliminar.Visible = False
        Me.cmdGrabar.Visible = False
        Me.cmdNuevo.Visible = False
        Me.cmdModificar.Visible = False
        Me.cmdCancelarPlanilla.Visible = False
        Me.cmdProcesar.Visible = False
        Me.cmdAbonarCuentas.Visible = False
        Me.cmdAsiento.Visible = False
        Me.cmdValida.Visible = False
        cmdexportar.Enabled = True
    End If
    
    Set Progress = New clsProgressBar
    
    If lsPlanillaCodDefecto <> "" Then
        Me.TxtPlanillas.Text = lsPlanillaCodDefecto
        txtPlanillas_EmiteDatos
    End If
End Sub

Private Sub IniFlex(pbIni As Boolean)

    Dim rsF As ADODB.Recordset
    Dim rsC As ADODB.Recordset
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    Set rsF = New ADODB.Recordset
    Set rsC = New ADODB.Recordset
    
    Flex.Rows = 1
    Flex.Rows = 2
    Flex.Cols = 5
    Flex.FixedRows = 1
    Flex.FixedCols = 4
    
    Flex.TextMatrix(0, 0) = "B"
    Flex.TextMatrix(0, 1) = "Cod.Emp."
    Flex.TextMatrix(0, 2) = "CodPers"
    Flex.TextMatrix(0, 3) = "Nombre"
    Flex.TextMatrix(0, 4) = "OK"
    
    Flex.ColWidth(0) = 1
    Flex.ColWidth(1) = 900
    Flex.ColWidth(2) = 1
    Flex.ColWidth(3) = 3900
    Flex.ColWidth(4) = 400
    
    Set rsF = oPla.GetPlanillasConceptos(Me.TxtPlanillas.Text)
    
    If Not RSVacio(rsF) Then
        While Not rsF.EOF
            Flex.Cols = Flex.Cols + 1
            Flex.TextMatrix(0, Flex.Cols - 1) = rsF!nemo
            If Left(rsF!Cab, 2) = RHConceptosTpoVVarUsuario Then
                Flex.ColWidth(Flex.Cols - 1) = 1
            Else
               Flex.ColWidth(Flex.Cols - 1) = Len(rsF!nemo) * 150
            End If
            rsF.MoveNext
        Wend
    End If
    
    Flex.Cols = Flex.Cols + 1
    Flex.ColWidth(Flex.Cols - 1) = 3500
    Flex.TextMatrix(0, Flex.Cols - 1) = "_Comentario"
    
    Set rsC = oPla.GetPlanillasPersona(Me.TxtPlanillas.Text, pbIni, Left(Me.txtPlanillaIns.Text, 8))
    
    If Not RSVacio(rsC) Then
        While Not rsC.EOF
            If Flex.TextMatrix(Flex.Rows - 1, 2) <> "" Then Flex.Rows = Flex.Rows + 1
            Flex.row = Flex.Rows - 1
            Flex.Col = 1
            Flex.CellBackColor = &HE0E0E0
            Flex.Col = 3
            Flex.CellBackColor = &HE0E0E0
            Flex.Col = 4
            'MAVM 20120214 FRACTAL ***
            'If TxtPlanillas.Text <> "E01" Then
            If TxtPlanillas.Text <> "22" Then
                Set Flex.CellPicture = picSi
                Flex.TextMatrix(Flex.Rows - 1, 0) = "1"
                Flex.TextMatrix(Flex.Rows - 1, 1) = rsC!cRhCod
            Else
                If Not pbIni Then
                    If Trim(rsC!nRHEstado) = 201 Then
                        Set Flex.CellPicture = picSi
                        Flex.TextMatrix(Flex.Rows - 1, 0) = "1"
                        Flex.TextMatrix(Flex.Rows - 1, 1) = rsC!cRhCod
                    Else
                        Set Flex.CellPicture = picNo
                        Flex.TextMatrix(Flex.Rows - 1, 0) = "0"
                        Flex.TextMatrix(Flex.Rows - 1, 1) = "_" & rsC!cRhCod
                    End If
                Else
                    If Not IsNull(rsC!nRHEstado) And Trim(rsC!nRHEstado) = 201 Or rsC!AgregaPlanilla Then
                        Set Flex.CellPicture = picSi
                        Flex.TextMatrix(Flex.Rows - 1, 0) = "1"
                        Flex.TextMatrix(Flex.Rows - 1, 1) = rsC!cRhCod
                    Else
                        Set Flex.CellPicture = Me.picNo
                        Flex.TextMatrix(Flex.Rows - 1, 0) = "0"
                        Flex.TextMatrix(Flex.Rows - 1, 1) = "_" & rsC!cRhCod
                        Flex.TextMatrix(Flex.Rows - 1, Flex.Cols - 1) = Trim(rsC!Estado)
                    End If
                End If
            End If
            Flex.CellPictureAlignment = 4
            Flex.TextMatrix(Flex.Rows - 1, 2) = rsC!codigo
            Flex.TextMatrix(Flex.Rows - 1, 3) = PstaNombre(rsC!nombre)
            rsC.MoveNext
        Wend
        ReDim Preserve lsPorMesTrab(Flex.Rows + 1)
    End If
    
    rsC.Close
    Set rsC = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = lnSalir
    If lnSalir <> 0 Then
        If MsgBox("Ud. no puede salir del formulario hasta que se acabe de procesar las planillas." & Chr(13) & "Desea Cancelar el proceso ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            lbCancela = True
            'ValidaProceso True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub

Private Sub mnuAgregar_Click()
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Dim lnI As Integer
    Set rsE = New ADODB.Recordset
    
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    
    If Not oPersona Is Nothing Then
        sqlE = oRRHH.GetCodigoEmpleado(oPersona.sPersCod)
        For lnI = 1 To Me.Flex.Rows - 1
            If Flex.TextMatrix(lnI, 1) = sqlE Then
                MsgBox "Empleado ya existe en la planilla.", vbInformation, "Aviso"
                Set rsE = Nothing
                Exit Sub
            End If
        Next lnI
        If Flex.TextMatrix(Flex.Rows - 1, 1) <> "" Then
            Flex.Rows = Flex.Rows + 1
        End If
        
        Flex.TextMatrix(Flex.Rows - 1, 0) = "1"
        Flex.TextMatrix(Flex.Rows - 1, 1) = sqlE
        Flex.TextMatrix(Flex.Rows - 1, 2) = oPersona.sPersCod
        Flex.TextMatrix(Flex.Rows - 1, 3) = PstaNombre(oPersona.sPersNombre, False)
        Flex.row = Flex.Rows - 1
        Flex.Col = 4
        Set Flex.CellPicture = picSi
        
    End If
End Sub

Private Sub mnuBuscar_Click()
    Dim i As Integer
    Dim lsCadena As String
    lsCadena = frmRHBuscarEmpleado.GetNombre(lsCadenaBuscar)
    lsCadenaBuscar = lsCadena
    For i = 1 To Flex.Rows - 1
        If InStr(1, Flex.TextMatrix(i, 3), lsCadena, vbTextCompare) <> 0 Then
            Me.Flex.TopRow = i
            Flex.row = i
            i = Flex.Rows - 1
        End If
    Next i
End Sub

Private Sub mnuBuscarSiguiente_Click()
    Dim i As Integer
    For i = Flex.row + 1 To Flex.Rows - 1
        If InStr(1, Flex.TextMatrix(i, 3), lsCadenaBuscar, vbTextCompare) <> 0 Then
            Me.Flex.TopRow = i
            Flex.row = i
            i = Flex.Rows - 1
        End If
    Next i
End Sub

Private Sub mnuComentario_Click()
    If Trim(Flex.TextMatrix(Flex.row, 1)) <> "" Then
        Flex.TextMatrix(Flex.row, Flex.Cols - 1) = frmRHPlanillaComentario.Ini(Me.Flex.TextMatrix(Flex.row, 1), Me.Flex.TextMatrix(Flex.row, 3), Flex.TextMatrix(Flex.row, Flex.Cols - 1), Me.TxtPlanillas.Text, Left(Me.txtPlanillaIns.Text, 8))
    End If
End Sub

Private Sub mskFecImp_GotFocus()
    mskFecImp.SelStart = 0
    mskFecImp.SelLength = 50
End Sub

Private Sub oPlaEvento_CloseProgress()
    Progress.CloseForm Me
End Sub

Private Sub oPlaEvento_Progress(pnValor As Long, pnTotal As Long)
    Progress.Max = pnTotal
    Progress.Progress pnValor, "Generando Reporte"
End Sub

Private Sub oPlaEvento_ShowProgress()
    Progress.ShowForm Me
End Sub

Private Sub TxtDes_GotFocus()
    txtDes.SelStart = 0
    txtDes.SelLength = 300
End Sub

Private Sub TxtDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecIni.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub TxtFecFin_GotFocus()
    txtFecFin.SelStart = 0
    txtFecFin.SelLength = 11
End Sub

Private Sub txtFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.Flex.SetFocus
End Sub

Private Sub TxtFecIni_GotFocus()
    txtFecIni.SelStart = 0
    txtFecIni.SelLength = 11
End Sub

Private Sub txtFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtFecFin.SetFocus
End Sub

Private Function Valida() As Boolean
    If Trim(TxtPlanillas.Text) = "" Then
        MsgBox "Debe indicar que tipo de Planilla se va a Generar.", vbInformation, "Aviso"
        TxtPlanillas.SetFocus
        Valida = False
    ElseIf Not txtFecIni.Enabled And TxtPlanillas <> gsRHPlanillaSubsidio Then
        MsgBox "Antes de Procesar debe elegir una nueva Planilla.", vbInformation, "Aviso"
        Valida = False
    ElseIf Not IsDate(txtFecIni) Then
        MsgBox "Debe indicar una fecha de Inicio de Proceso Valida.", vbInformation, "Aviso"
        txtFecIni.SetFocus
        Valida = False
    ElseIf Not IsDate(txtFecFin) Then
        MsgBox "Debe indicar una fecha de Fin de Proceso Valida.", vbInformation, "Aviso"
        txtFecFin.SetFocus
        Valida = False
    ElseIf Me.txtDes.Text = "" Then
        MsgBox "Debe indicar un comentario valido.", vbInformation, "Aviso"
        txtDes.SetFocus
        Valida = False
    Else
        Valida = True
    End If
End Function

Private Sub ClearScreen(Optional pbTotal As Boolean = True)
    txtDes = ""
    txtFecFin.Mask = ""
    txtFecIni.Mask = ""
    txtFecFin.Text = ""
    txtFecIni.Text = ""
    txtFecFin.Mask = "##/##/####"
    txtFecIni.Mask = "##/##/####"
    cmbEstadoPla.ListIndex = -1
    txtPlanillaIns.Text = ""
    If pbTotal Then TxtPlanillas.Text = ""
End Sub

Private Sub Activa(Optional pbTotal As Boolean = True, Optional pbvalor As Boolean = True)
    If pbTotal Then
        Me.fraPla.Enabled = pbvalor
    Else
        txtPlanillaIns.Enabled = Not pbvalor
        txtDes.Enabled = pbvalor
        txtFecFin.Enabled = pbvalor
        txtFecIni.Enabled = pbvalor
        cmbEstadoPla.Enabled = pbvalor
    End If
End Sub

Private Sub GetPosIJ(psCodEmp As String, psCodCon As String, pnI As Integer, pnJ As Integer)
    Dim i As Integer
    
    For i = 1 To Flex.Rows - 1
        'MAVM 20120403 ***
        'If Right(Flex.TextMatrix(i, 1), 6) = psCodEmp Then
        If Flex.TextMatrix(i, 1) = psCodEmp Then
            pnI = i
            i = Flex.Rows - 1
        End If
    Next i
    
    For i = 5 To Flex.Cols - 1
        If Trim(Flex.TextMatrix(0, i)) = psCodCon Then
            pnJ = i
            i = Flex.Cols - 1
        End If
    Next i
End Sub

'Private Function VerPlaPagEmp(psCodEmp As String, psCodPla As String, pdFecha As Date) As Boolean
'    Dim sqlE As String
'    Dim rsE As New ADODB.Recordset
'
'    sqlE = " Select cPlaCod From Planilladetalle" _
'         & " Where cPlaCod = '" & psCodPla & "' And cEmpCod = '" & psCodEmp & "' And dPlaInsCod = '" & Format(pdFecha, gsformatofechahora) & "' And bPagado = 1"
'    rsE.Open sqlE, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If Not RSVacio(rsE) Then
'        VerPlaPagEmp = True
'    Else
'        VerPlaPagEmp = False
'    End If
'
'    RSCierra rsE
'End Function

Private Sub ValidaProceso(pbvalor As Boolean)
    TxtPlanillas.Enabled = pbvalor
    txtFecIni.Enabled = pbvalor
    txtFecFin.Enabled = pbvalor
    cmdProcesar.Enabled = pbvalor
    cmdAbonarCuentas.Enabled = pbvalor
    
    cmdCancelar.Visible = Not pbvalor
    cmdSalir.Visible = pbvalor
End Sub

Private Sub ValidaBotones(pbvalor As Boolean)
    cmdGrabar.Visible = pbvalor
    cmdCancelar.Visible = pbvalor
    cmdNuevo.Visible = Not pbvalor
    cmdModificar.Visible = Not pbvalor
    cmdImprimir.Enabled = Not pbvalor
    cmdEliminar.Enabled = Not pbvalor
    cmdAbonarCuentas.Enabled = pbvalor
    cmdProcesar.Enabled = pbvalor
End Sub

'Private Function GetEmpleadoEstado(psEmpCod As String) As String
'    Dim sqlC As String
'    Dim rsC As ADODB.Recordset
'    Set rsC = New ADODB.Recordset
'    sqlC = " Select TC.cNomTab from empleado E " _
'         & " Inner Join dbComunes..TablaCod TC On E.cEmpEst = TC.cValor and cCodTab like 'EB__'" _
'         & " where E.cEmpCod = '" & psEmpCod & "'"
'    rsC.Open sqlC, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'    If RSVacio(rsC) Then
'        GetEmpleadoEstado = ""
'    Else
'        GetEmpleadoEstado = Trim(rsC!cNomTab)
'    End If
'
'    rsC.Close
'    Set rsC = Nothing
'End Function

'******************************************************
Public Sub Ini(pnTipo As TipoProcesoRRHH, psCaption As String, pMdi As Form, Optional psPlanillaCodDefecto As String = "")
    lsPlanillaCodDefecto = psPlanillaCodDefecto
    lnTipo = pnTipo
    Caption = psCaption
    Me.Show , pMdi
End Sub

Private Sub txtPlanillaIns_EmiteDatos()
    Me.lblPlaniInstRes.Caption = txtPlanillaIns.psDescripcion
    lsPlanillaAnt = Me.txtPlanillaIns
    CargaPlanillaExistente
    'Me.Flex.SetFocus
End Sub

Private Sub txtPlanillaIns_OnValidaClick(Vacio As Boolean)
    If lsPlanillaAnt = Me.txtPlanillaIns Then Vacio = True
End Sub

Private Sub txtPlanillaIns_Validate(Cancel As Boolean)
    lsPlanillaAnt = Me.txtPlanillaIns
End Sub

Private Sub txtPlanillas_EmiteDatos()
    Me.lblPlanillaRes.Caption = TxtPlanillas.psDescripcion
    If TxtPlanillas.Text = "" Then Exit Sub
    CargaPlanillasTpo TxtPlanillas.Text
    txtPlanillaIns_EmiteDatos
    Dim oDHCierre As DRHProcesosCierre
    Set oDHCierre = New DRHProcesosCierre
    lsOpeCodFractal = oDHCierre.GetObtenerCodigoOperacion(Trim(TxtPlanillas.Text))
End Sub

Private Sub CargaPlanillasTpo(psPlaCodigo As String)
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    
    If TxtPlanillas.Text <> "" And Not fraPla.Enabled Then fraPla.Enabled = True
    
    Set rsP = oPla.GetPlanillasTpo(psPlaCodigo)
    Me.txtPlanillaIns.rs = rsP
    
    IniFlex False
End Sub

Private Sub CargaPlanillaExistente()
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Dim i As Integer
    Dim J As Integer
    Dim lnAcumulador As Currency
    Dim lnAcumuladorDolares As Currency 'MAVM 20120505
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    
    IniFlex False
    
    Set rsP = oPla.GetPlanillasExistenteDatos(Me.TxtPlanillas.Text, Left(Me.txtPlanillaIns.Text, 8))
    
    If Not RSVacio(rsP) Then
        txtDes = Trim(rsP!Des)
        txtFecIni = Format(rsP!dIni, gsFormatoFechaView)
        txtFecFin = Format(rsP!dFin, gsFormatoFechaView)
        Me.mskFecImp.Text = Format(rsP!Impre, gsFormatoFechaView)
        Me.txtTpoCambio.Text = Format(rsP!TC, "#00.000")
        Me.txtCambFij.Text = Format(rsP!TCf, "#00.000")
        UbicaCombo cmbEstadoPla, Trim(rsP!Estado)
        
        If rsP!Estado = "2" Then
            cmdAbonarCuentas.Enabled = False
        Else
            cmdAbonarCuentas.Enabled = True
        End If
    End If
    
    Set rsP = oPla.GetPlanillasExistenteDetalle(Me.TxtPlanillas.Text, Left(Me.txtPlanillaIns.Text, 8))
    
    If rsP Is Nothing Then Exit Sub
    If Not RSVacio(rsP) Then
        While Not rsP.EOF
            i = -1
            J = -1
            GetPosIJ rsP!cRhCod, Trim(rsP!cRHconceptoMeno), i, J
            'MAVM 20120505 ***
            'If i <> -1 And j <> -1 Then Flex.TextMatrix(i, j) = Format(rsP!nMonto, "#,##0.00")
            If i <> -1 And J <> -1 Then
                Flex.TextMatrix(i, J) = Format(rsP!nMonto, "#,##0.00")
                Flex.TextMatrix(i, J + 1) = rsP!cCtaCod
            End If
            '***
            If rsP!Orden = 9999 Then
                Flex.row = i
                Flex.Col = J
                Flex.CellBackColor = &HC0C000
                Flex.CellFontBold = True
            End If
            rsP.MoveNext
        Wend
    End If
    
    Set rsP = oPla.GetPlanillasExistenteComentario(Me.TxtPlanillas.Text, Left(Me.txtPlanillaIns.Text, 8))
    
    If Not RSVacio(rsP) Then
        While Not rsP.EOF
            GetPosIJ rsP!cRhCod, Trim(rsP!cConcepCod), i, J
            Flex.TextMatrix(i, J) = IIf(IsNull(rsP!cComentario), "", rsP!cComentario)
            rsP.MoveNext
        Wend
    End If
    
    rsP.Close
    Set rsP = Nothing
    
    If Flex.TextMatrix(Flex.Rows - 1, 2) <> "" Then
        Flex.Rows = Flex.Rows + 1
        Flex.TextMatrix(Flex.Rows - 1, 3) = "TOTAL"
        
        For J = 5 To Me.Flex.Cols - 1
            lnAcumulador = 0
            lnAcumuladorDolares = 0
            If Left(Flex.TextMatrix(0, J), 2) <> "U_" And Left(Flex.TextMatrix(0, J), 1) <> "_" Then
                
                For i = 1 To Me.Flex.Rows - 2
                    If Flex.TextMatrix(i, J) <> "" Then
                        If Mid(Flex.TextMatrix(i, J + 1), 9, 1) = "1" Then
                            lnAcumulador = lnAcumulador + CCur(Flex.TextMatrix(i, J))
                        Else
                            lnAcumuladorDolares = lnAcumuladorDolares + CCur(Flex.TextMatrix(i, J))
                        End If
                    End If
                Next i
                'Flex.TextMatrix(Flex.Rows - 1, j) = Format(lnAcumulador, "#,##.00")
                Flex.row = Flex.Rows - 1
                Flex.Col = J
                Flex.CellBackColor = &HA0C000
                Flex.CellFontBold = True
                'lnAcumulador = lnAcumulador + CCur(Flex.TextMatrix(i, j))
                lblSol.Caption = Format(lnAcumulador, "#,##.00")
                lblDol.Caption = Format(lnAcumuladorDolares, "#,##.00")
            End If
        Next J
    End If
    Flex.TextMatrix(Flex.Rows - 1, 1) = Flex.Rows - 2
End Sub
