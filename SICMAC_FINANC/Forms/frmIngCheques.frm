VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIngCheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Cheques"
   ClientHeight    =   6660
   ClientLeft      =   1875
   ClientTop       =   1590
   ClientWidth     =   8295
   Icon            =   "frmIngCheques.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraIngCheque 
      Caption         =   "Datos Generales"
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
      Height          =   1770
      Left            =   90
      TabIndex        =   24
      Top             =   105
      Width           =   8115
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   7110
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1365
         Width           =   885
      End
      Begin Sicmact.TxtBuscar txtIngChqBuscaIF 
         Height          =   360
         Left            =   1095
         TabIndex        =   0
         Top             =   270
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   635
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
      Begin VB.ComboBox cboIngChqPlaza 
         Height          =   315
         ItemData        =   "frmIngCheques.frx":030A
         Left            =   945
         List            =   "frmIngCheques.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1305
         Width           =   1380
      End
      Begin VB.TextBox txtIngChqNumCheque 
         Height          =   300
         Left            =   4035
         MaxLength       =   15
         TabIndex        =   2
         Top             =   945
         Width           =   1605
      End
      Begin VB.TextBox txtIngChqCtaIF 
         Height          =   315
         Left            =   945
         MaxLength       =   20
         TabIndex        =   1
         Top             =   908
         Width           =   2085
      End
      Begin VB.CheckBox chkConfirmar 
         Caption         =   "Por Confimar en Caja Gen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5820
         TabIndex        =   3
         Top             =   975
         Width           =   2160
      End
      Begin MSMask.MaskEdBox txtIngChqFechaReg 
         Height          =   315
         Left            =   3090
         TabIndex        =   5
         Top             =   1335
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtIngChqFechaVal 
         Height          =   315
         Left            =   5250
         TabIndex        =   6
         Top             =   1350
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         Height          =   195
         Left            =   6435
         TabIndex        =   36
         Top             =   1395
         Width           =   630
      End
      Begin VB.Label lblEstado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   225
         Left            =   5895
         TabIndex        =   35
         Top             =   660
         Width           =   2040
      End
      Begin VB.Label lblIngChqDescIF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3225
         TabIndex        =   32
         Top             =   285
         Width           =   4635
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Cheque"
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
         Height          =   210
         Left            =   120
         TabIndex        =   31
         Top             =   645
         Width           =   645
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Valorización:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4305
         TabIndex        =   30
         Top             =   1395
         Width           =   945
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Plaza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   29
         Top             =   1357
         Width           =   390
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Registro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2445
         TabIndex        =   28
         Top             =   1365
         Width           =   645
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "N° Cheque :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3135
         TabIndex        =   27
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Institución Financiera :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   26
         Top             =   210
         Width           =   900
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta N° :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   25
         Top             =   960
         Width           =   810
      End
   End
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   8115
      Begin VB.TextBox txtMovDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   180
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   210
         Width           =   7710
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   5820
         TabIndex        =   13
         Top             =   1020
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "Monto  : "
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
         Left            =   4980
         TabIndex        =   17
         Top             =   1065
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   6960
      TabIndex        =   15
      Top             =   6240
      Width           =   1290
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   5685
      TabIndex        =   14
      Top             =   6240
      Width           =   1290
   End
   Begin VB.Frame fraEspecifica 
      Height          =   2760
      Left            =   105
      TabIndex        =   18
      Top             =   1890
      Width           =   8115
      Begin VB.Frame Frame2 
         Caption         =   "Lugar de Recepción"
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
         Height          =   1080
         Left            =   90
         TabIndex        =   19
         Top             =   180
         Width           =   7860
         Begin Sicmact.TxtBuscar txtBuscarProd 
            Height          =   330
            Left            =   1170
            TabIndex        =   8
            Top             =   285
            Width           =   1395
            _ExtentX        =   2461
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
         Begin Sicmact.TxtBuscar txtBuscarAreaAgencia 
            Height          =   330
            Left            =   1170
            TabIndex        =   9
            Top             =   630
            Width           =   1395
            _ExtentX        =   2461
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
         Begin VB.Label Label2 
            Caption         =   "Producto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   23
            Top             =   300
            Width           =   645
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Area/Agencia:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   105
            TabIndex        =   22
            Top             =   690
            Width           =   1050
         End
         Begin VB.Label lblProdDesc 
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
            Height          =   315
            Left            =   2595
            TabIndex        =   21
            Top             =   285
            Width           =   5190
         End
         Begin VB.Label lblAreaAgeDesc 
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
            Height          =   315
            Left            =   2595
            TabIndex        =   20
            Top             =   645
            Width           =   5175
         End
      End
      Begin Sicmact.FlexEdit fgObjMotivo 
         Height          =   990
         Left            =   255
         TabIndex        =   11
         Top             =   1680
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   1746
         Cols0           =   5
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "-Objeto-Descripción-SubCta-cObjetoCod"
         EncabezadosAnchos=   "350-1600-4000-1000-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   285
      End
      Begin Sicmact.TxtBuscar txtBuscarCtaHaber 
         Height          =   330
         Left            =   1245
         TabIndex        =   10
         Top             =   1290
         Width           =   1395
         _ExtentX        =   2461
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Motivo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   330
         TabIndex        =   34
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label lblCtaHaber 
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
         Left            =   2655
         TabIndex        =   33
         Top             =   1305
         Width           =   5160
      End
   End
End
Attribute VB_Name = "frmIngCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCtasIF As NCajaCtaIF
Dim oContFunct As NContFunciones
Dim oGen As DGeneral
Dim oOpe As DOperacion
Dim oCtaCont As DCtaCont

Dim rsPend As ADODB.Recordset

Dim lbOk As Boolean
Dim lsPersCodIf  As String
Dim lsNroCtaIf As String
Dim lsNroChq As String
Dim lnPlazaChq As ChequePlaza
Dim ldFechaRegChq As Date
Dim ldFechaValChq As Date
Dim lsConfCheque As String
Dim lsGlosa As String
Dim lnMonto As Currency
Dim lnImporte As Double

Dim lsCtaContChq As String
Dim lbMuestra As Boolean
Dim lbNegocio As Boolean
Dim lsOpeCod As String
Dim lsCtaContHaber As String
Dim lnDiasValoriza As Long
Dim lbRegCheque As Boolean
Dim lsMovRef As String
Dim lsNombreIF As String

'variables para Arendir
Dim lsMovNroAtenc As String
Dim lsMovNroSol As String
Dim lnTipoArendir As ArendirTipo
Dim lbArendir As Boolean
Dim lnMoneda  As Moneda
Dim lnOrdenProd As String
Dim lnOrdenAgencia As String
Dim lbApertura As Boolean
Dim lbRendirPendiente As Boolean

Dim lbCreaSubCta As Boolean
Dim lsSubCtaIFCod As String
Dim lsSubCtaIFDesc As String
Dim lsPersCodAper As String
Dim lsIFTpoAper As String
Dim lsCtaIFCod As String
Dim lsCtaIFDesc As String
Dim ldCtaIFAper As Date
Dim lsCtaIFVenc As String
Dim lnCtaIFPlazo As Integer
Dim lnPeriodo As Integer
Dim lnInteres  As Currency
Dim lnTpoDocAper As TpoDoc
Dim lsNroDocAper As String
Dim ldFechaDocApera As Date
Dim lsDocumentoAper As String

Dim rsObj As ADODB.Recordset
Dim lbSoloIngreso  As Boolean
Dim lsProductoCod As String
Dim lsAreaAgeCod As String
Dim lsCtaMotivo As String

Private Sub CboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If fraEspecifica.Visible Then
        If txtBuscarProd.Enabled Then
            txtBuscarProd.SetFocus
        Else
            If txtBuscarCtaHaber.Enabled Then
                txtBuscarCtaHaber.SetFocus
            Else
                txtMovDesc.SetFocus
            End If
        End If
    Else
        txtMovDesc.SetFocus
    End If
End If
End Sub

Private Sub chkConfirmar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboIngChqPlaza.SetFocus
End If
End Sub
Private Sub cmdAceptar_Click()
Dim oDocRec As NDocRec
Dim oCaja As nCajaGeneral

Dim lsMovNro As String
Dim oArendir As NARendir
Dim rs As ADODB.Recordset
Dim lnConfCaja As CGEstadoConfCheque
Dim lsCtaDebe As String

Set oCaja = New nCajaGeneral
Set rs = New ADODB.Recordset
If lbMuestra = False Then
    If ValidaInterfaz = False Then Exit Sub
    
    If lbSoloIngreso = False Then
        Set oDocRec = New NDocRec
        Set oArendir = New NARendir
        If MsgBox("Desea Realizar el Registro del Cheque??", vbYesNo + vbInformation, "Aviso") = vbYes Then
            lsMovNro = oContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            If lbNegocio = True Then
                If chkConfirmar.value Then
                    lnConfCaja = ChqCGNoConfirmado
                Else
                    lnConfCaja = ChqCGSinConfirmacion
                End If
                oDocRec.RegistroChequesNegocio lsMovNro, lsOpeCod, txtMovDesc, Mid(txtIngChqBuscaIF, 1, 2), txtIngChqNumCheque, _
                                                Mid(txtIngChqBuscaIF, 4, 13), cboIngChqPlaza.ListIndex, txtIngChqCtaIF, _
                                                CCur(txtMonto), CDate(txtIngChqFechaReg), CDate(txtIngChqFechaVal), gsFormatoFecha, Right(cboMoneda, 1), , , lnConfCaja, gsCodArea, gsCodAge
            Else
                If lbArendir = True Then
                    'edpyme
                    lsCtaDebe = lsCtaDebe + oContFunct.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaContChq, txtBuscarAreaAgencia, False, False)
                    lsCtaContChq = lsCtaContChq + lsCtaDebe
                    
                    oArendir.GrabaRendicionIngresoCheque lnTipoArendir, gsFormatoFecha, lsMovNro, lsOpeCod, txtMovDesc, _
                                lsCtaContHaber, lsCtaContChq, txtBuscarProd.Text, Mid(txtBuscarAreaAgencia.Text, 1, 3), Mid(txtBuscarAreaAgencia.Text, 4, 2), _
                                CCur(txtMonto), lsMovNroAtenc, lsMovNroSol, Trim(txtIngChqNumCheque), _
                                Mid(txtIngChqBuscaIF, 4, 13), Mid(txtIngChqBuscaIF, 1, 2), cboIngChqPlaza.ListIndex, txtIngChqCtaIF, CDate(txtIngChqFechaReg), _
                                CDate(txtIngChqFechaVal), Right(cboMoneda, 1), gChqEstEnValorizacion, gCGEstadosChqRecibido, ChqCGConfirmado, Mid(txtBuscarAreaAgencia, 1, 3), Mid(txtBuscarAreaAgencia, 4, 2)
                                         
                ElseIf lbRendirPendiente = True Then
                    Dim oAna As New NAnalisisCtas
                    oAna.GrabaRendirPendIngresoCheque lsMovNro, lsOpeCod, txtMovDesc, _
                                lsCtaContHaber, lsCtaContChq, txtBuscarProd.Text, Mid(txtBuscarAreaAgencia.Text, 1, 3), Mid(txtBuscarAreaAgencia.Text, 4, 2), _
                                CCur(txtMonto), Trim(txtIngChqNumCheque), _
                                Mid(txtIngChqBuscaIF, 4, 13), Mid(txtIngChqBuscaIF, 1, 2), cboIngChqPlaza.ListIndex, txtIngChqCtaIF, CDate(txtIngChqFechaReg), _
                                CDate(txtIngChqFechaVal), Right(cboMoneda, 1), rsPend, gChqEstEnValorizacion, gCGEstadosChqRecibido, ChqCGConfirmado, Mid(txtBuscarAreaAgencia, 1, 3), Mid(txtBuscarAreaAgencia, 4, 2)
                Else
                    If fgObjMotivo.TextMatrix(1, 0) <> "" Then
                        Set rs = fgObjMotivo.GetRsNew
                    End If
                    
                    lsCtaDebe = oContFunct.GetFiltroObjetos(ObjProductosCMACT, lsCtaContChq, txtBuscarProd, False)
                    'John agregue
                    lsCtaDebe = lsCtaDebe + oContFunct.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaContChq, txtBuscarAreaAgencia, False)
                    '**********
                    lsCtaDebe = lsCtaContChq + lsCtaDebe
                    
                    If lbApertura = False Then
                        oDocRec.RegistroChequesContab lsMovNro, lsOpeCod, txtMovDesc, Val(lsMovRef), lsCtaDebe, txtBuscarProd, _
                                    Mid(txtBuscarAreaAgencia.Text, 1, 3), Mid(txtBuscarAreaAgencia.Text, 4, 2), txtBuscarCtaHaber.Text, _
                                    rs, txtIngChqNumCheque, Mid(txtIngChqBuscaIF, 4, 13), Mid(txtIngChqBuscaIF, 1, 2), cboIngChqPlaza.ListIndex, _
                                    Me.txtIngChqCtaIF, CCur(txtMonto), CDate(txtIngChqFechaReg), CDate(txtIngChqFechaVal), _
                                    gsFormatoFecha, Right(cboMoneda, 1), gChqEstValorizado, gCGEstadosChqRecibido, ChqCGConfirmado, Mid(txtBuscarAreaAgencia, 1, 3), Mid(txtBuscarAreaAgencia, 4, 2)
                    Else
                        
                        oCaja.GrabaAperturaRegCheque lsMovNro, lsOpeCod, txtMovDesc, txtBuscarCtaHaber.Text, txtBuscarProd, _
                                    Mid(txtBuscarAreaAgencia.Text, 1, 3), Mid(txtBuscarAreaAgencia.Text, 4, 2), lsCtaDebe, _
                                    rs, txtIngChqNumCheque, Mid(txtIngChqBuscaIF, 4, 13), Mid(txtIngChqBuscaIF, 1, 2), cboIngChqPlaza.ListIndex, _
                                    txtIngChqCtaIF, CCur(txtMonto), CDate(txtIngChqFechaReg), CDate(txtIngChqFechaVal), _
                                    Right(cboMoneda, 1), _
                                    lbCreaSubCta, lsSubCtaIFCod, lsSubCtaIFDesc, lsPersCodAper, _
                                    lsIFTpoAper, lsCtaIFCod, lsCtaIFDesc, ldCtaIFAper, _
                                    lsCtaIFVenc, lnCtaIFPlazo, lnPeriodo, lnInteres, lnTpoDocAper, lsNroDocAper, _
                                    ldFechaDocApera, _
                                    gChqEstValorizado, gCGEstadosChqRecibido, ChqCGConfirmado, Mid(txtBuscarAreaAgencia, 1, 3), Mid(txtBuscarAreaAgencia, 4, 2)
                                    
                        If lsDocumentoAper <> "" Then
                                EnviaPrevio lsDocumentoAper & oImpresora.gPrnSaltoPagina & lsDocumentoAper, "Carta Apertura", gnLinPage, False
                        End If
                    End If
                End If
                ImprimeAsientoContable lsMovNro, "", "", ""
            End If
            EmiteDatos
            Set oDocRec = Nothing
            Set oArendir = Nothing
            lbOk = True
            Unload Me
        End If
    Else
        EmiteDatos
        lbOk = True
        Unload Me
    End If
Else
    EmiteDatos
    lbOk = True
    Unload Me
End If
End Sub
Private Sub cmdCancelar_Click()
lbOk = False
Unload Me
End Sub
Private Sub Form_Load()
Dim oCapDef As nCapDefinicion

Set oCtaCont = New DCtaCont
Set oCtasIF = New NCajaCtaIF
Set oContFunct = New NContFunciones
Set oGen = New DGeneral
Set oOpe = New DOperacion
Set oCapDef = New nCapDefinicion
lnDiasValoriza = oCapDef.GetCapParametro(gDiasMinValChq)
Set oCapDef = Nothing
'datos para Ing de cheques
CambiaTamañoCombo cboIngChqPlaza, 150
CambiaTamañoCombo cboMoneda, 100
txtIngChqFechaReg = gdFecSis
txtIngChqFechaVal = gdFecSis
Me.Caption = " Registro de Cheques - " & gsOpeDesc
 If lbApertura Then
    lsCtaContChq = oOpe.EmiteOpeCta(lsOpeCod, "H", "1")
Else
    lsCtaContChq = oOpe.EmiteOpeCta(lsOpeCod, "D")
End If
cboMoneda.Enabled = False
txtMonto = Format(lnMonto, "#,#0.00")
txtMonto.Locked = lbMuestra
If lnMonto > 0 Then
    txtMonto.Locked = True
End If

txtIngChqBuscaIF.rs = oCtasIF.GetInstFinancieras("_" & gTpoIFBanco & "%")
FraIngCheque.Enabled = Not lbMuestra
txtMovDesc.Locked = lbMuestra
txtMovDesc = lsGlosa

CargaCombo cboMoneda, oGen.GetConstante(gMoneda)
If Mid(gsOpeCod, 3, 1) = "2" Then
    cboMoneda.ListIndex = 1
Else
    cboMoneda.ListIndex = 0
End If

txtBuscarProd.rs = oOpe.GetOpeObj(lsOpeCod, lnOrdenProd)
If txtBuscarProd.rs.State = adStateOpen Then
    If txtBuscarProd.rs.RecordCount = 1 Then
        txtBuscarProd.Text = txtBuscarProd.rs(0)
        lblProdDesc = txtBuscarProd.psDescripcion
        txtBuscarProd.Enabled = False
    End If
End If

txtBuscarAreaAgencia.rs = oOpe.GetOpeObj(lsOpeCod, lnOrdenAgencia)
If txtBuscarAreaAgencia.rs.State = adStateOpen Then
    If txtBuscarAreaAgencia.rs.RecordCount = 1 Then
        txtBuscarAreaAgencia.Text = txtBuscarAreaAgencia.rs(0)
        lblAreaAgeDesc = txtBuscarAreaAgencia.psDescripcion
        txtBuscarAreaAgencia.Enabled = False
    End If
End If
If lsCtaContHaber <> "" Then
    txtBuscarCtaHaber = lsCtaContHaber
    lblCtaHaber = oGen.CuentaNombre(lsCtaContHaber)
    txtBuscarCtaHaber.Enabled = False
Else
    txtBuscarCtaHaber.psRaiz = "Cuentas Contables"
    txtBuscarCtaHaber.rs = oOpe.CargaOpeCta(lsOpeCod, "H", "0")
End If
If lbNegocio = False Then
    fraGlosa.Top = fraEspecifica.Top + fraEspecifica.Height + 50
    cmdAceptar.Top = fraGlosa.Top + fraGlosa.Height + 50 '6000
    cmdCancelar.Top = fraGlosa.Top + fraGlosa.Height + 50 '6000
    Me.fraEspecifica.Visible = True
    Me.Height = cmdCancelar.Top + cmdCancelar.Height + 500
    Me.chkConfirmar.Visible = False
Else
    fraGlosa.Top = FraIngCheque.Top + FraIngCheque.Height + 100
    cmdAceptar.Top = fraGlosa.Top + fraGlosa.Height + 50 '6000
    cmdCancelar.Top = fraGlosa.Top + fraGlosa.Height + 50 '6000
    fraEspecifica.Visible = False
    Me.Height = cmdCancelar.Top + cmdCancelar.Height + 500
    txtIngChqFechaReg.Enabled = False
    txtIngChqFechaVal = DateAdd("d", lnDiasValoriza, CDate(txtIngChqFechaReg))
End If
AsignaDatos
CentraForm Me
End Sub
Public Sub InicioMuestra(ByVal psPersCodIF As String, ByVal psNroChq As String, _
                        ByVal pbNegocio As Boolean, ByVal psOpeCod As String)

lbMuestra = True
lbArendir = False
lsPersCodIf = psPersCodIF
lbApertura = False
lsNroChq = psNroChq
lsOpeCod = psOpeCod
lbNegocio = pbNegocio
Me.Show 1
End Sub
Public Sub Inicio(ByVal pbNegocio As Boolean, ByVal psOpeCod As String, _
                  ByVal pbRegCheque As Boolean, ByVal pnMonto As Currency, ByVal pnMoneda As Moneda, _
                  Optional ByVal psMovref As String = "", _
                  Optional pnOrdenProd As Integer = 1, Optional pnOrdenAgencia As Integer = 2, Optional pbSoloIngresa As Boolean = False, Optional psGlosa As String = "")
lbMuestra = False
lbArendir = False
lbNegocio = pbNegocio
lnMoneda = pnMoneda
lsOpeCod = psOpeCod
lbRegCheque = pbRegCheque
lbRendirPendiente = False
lsMovRef = psMovref
lnMonto = pnMonto
lnOrdenProd = pnOrdenProd
lnOrdenAgencia = pnOrdenAgencia
lbApertura = False
lbSoloIngreso = pbSoloIngresa
lsGlosa = psGlosa
Me.Show 1
End Sub
Public Sub InicioArendir(ByVal psOpeCod As String, ByVal pnMonto As Currency, ByVal pnTipoArendir As ArendirTipo, _
                        ByVal psMovNroAtenc As String, ByVal psMovNroSol As String, ByVal psCtaContPendiente As String, ByVal psGlosa As String, ByVal pnMoneda As Moneda, _
                        Optional pnOrdenProd As Integer = 1, Optional pnOrdenAgencia As Integer = 2)
lbMuestra = False
lbArendir = True
lbRendirPendiente = False
lsOpeCod = psOpeCod
lnMoneda = pnMoneda
lsMovNroAtenc = psMovNroAtenc
lsMovNroSol = psMovNroSol
lnTipoArendir = pnTipoArendir
lsCtaContHaber = psCtaContPendiente
lsGlosa = psGlosa
lnMonto = pnMonto
lnOrdenProd = pnOrdenProd
lnOrdenAgencia = pnOrdenAgencia
lbApertura = False
Me.Show 1
End Sub

Public Sub InicioRendirPendiente(ByVal psOpeCod As String, ByVal pnMonto As Currency, _
                        ByVal psCtaContPendiente As String, ByVal psGlosa As String, ByVal pnMoneda As Moneda, _
                        ByVal prsPend As ADODB.Recordset, _
                        Optional pnOrdenProd As Integer = 1, Optional pnOrdenAgencia As Integer = 2)
lbMuestra = False
lbArendir = False
lbRendirPendiente = True
lsOpeCod = psOpeCod
lnMoneda = pnMoneda
lsCtaContHaber = psCtaContPendiente
lsGlosa = psGlosa
lnMonto = pnMonto
lnOrdenProd = pnOrdenProd
lnOrdenAgencia = pnOrdenAgencia
lbApertura = False
Set rsPend = prsPend
Me.Show 1
End Sub


Public Sub InicioAperturas(ByVal psOpeCod As String, ByVal pnMonto As Currency, _
                          ByVal psCtaHaber As String, ByVal psGlosa As String, ByVal pnMoneda As Moneda, _
                          ByVal pbCreaSubCta As Boolean, ByVal psSubCtaIFCod As String, ByVal psSubCtaIFDesc As String, _
                          ByVal psPersCod As String, ByVal psIFTpo As String, _
                          ByVal psCtaIFCod As String, ByVal psCtaIFDesc As String, _
                          ByVal pdCtaIFAper As Date, ByVal psCtaIFVenc As String, _
                          ByVal pnCtaIFPlazo As Integer, ByVal pnPeriodo As Integer, ByVal pnInteres As Currency, _
                          ByVal pnTpoDocAper As TpoDoc, ByVal psNroDocAper As String, _
                          ByVal pdFechaDocApera As String, ByVal psDocumentoAper As String, _
                          Optional pnOrdenProd As Integer = 1, Optional pnOrdenAgencia As Integer = 2)


lbCreaSubCta = pbCreaSubCta
lsSubCtaIFCod = psSubCtaIFCod
lsSubCtaIFDesc = psSubCtaIFDesc
lsPersCodAper = psPersCod
lsIFTpoAper = psIFTpo
lsCtaIFCod = psCtaIFCod
lsCtaIFDesc = psCtaIFDesc
ldCtaIFAper = pdCtaIFAper
lsCtaIFVenc = psCtaIFVenc
lnCtaIFPlazo = pnCtaIFPlazo
lnTpoDocAper = pnTpoDocAper
lsNroDocAper = psNroDocAper
ldFechaDocApera = pdFechaDocApera
lnTpoDocAper = pnTpoDocAper
lsNroDocAper = psNroDocAper
ldFechaDocApera = pdFechaDocApera
lsDocumentoAper = psDocumentoAper
lnPeriodo = pnPeriodo
lnInteres = pnInteres

lbMuestra = False
lbRendirPendiente = False
lbApertura = True
lsOpeCod = psOpeCod
lnMoneda = pnMoneda
lsCtaContHaber = psCtaHaber
lsGlosa = psGlosa
lnMonto = pnMonto
lnOrdenProd = pnOrdenProd
lnOrdenAgencia = pnOrdenAgencia

Me.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oCtasIF = Nothing
    Set oContFunct = Nothing
    Set oGen = Nothing
    Set oOpe = Nothing
    Set oCtaCont = Nothing
End Sub
Private Function ValidaInterfaz() As Boolean
Dim oDoc As DDocumento

Set oDoc = New DDocumento

ValidaInterfaz = True
If Len(Trim(txtIngChqBuscaIF)) = 0 Or Len(Trim(lblIngChqDescIF)) = 0 Then
    MsgBox "Institución Financiera no válida", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtIngChqBuscaIF.SetFocus
    Exit Function
End If
If Len(Trim(txtIngChqCtaIF)) = 0 Then
    MsgBox "Nro de Cuenta de Institución Financiera no válida", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtIngChqCtaIF.SetFocus
    Exit Function
End If
If Len(Trim(txtIngChqNumCheque)) = 0 Then
    MsgBox "Nro de Cheque no Ingresado o no es válida", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtIngChqNumCheque.SetFocus
    Exit Function
End If
If oDoc.VerificaDoc(TpoDocCheque, txtIngChqNumCheque, Mid(txtIngChqBuscaIF, 4, 13)) Then
    MsgBox "Documento ya se encuentra registrado ", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtIngChqNumCheque.SetFocus
    Exit Function
End If
If oDoc.VerificaCheque(TpoDocCheque, txtIngChqNumCheque, Mid(txtIngChqBuscaIF, 4, 13), Mid(txtIngChqBuscaIF, 1, 2)) Then
    MsgBox "Cheque ya se encuentra registrado ", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtIngChqNumCheque.SetFocus
    Exit Function
End If


If cboIngChqPlaza = "" Then
    MsgBox "Plaza de Cheque no válido", vbInformation, "Aviso"
    ValidaInterfaz = False
    cboIngChqPlaza.SetFocus
    Exit Function
End If
If ValFecha(txtIngChqFechaReg) = False Then
    ValidaInterfaz = False
    Exit Function
End If
If ValFecha(txtIngChqFechaVal) = False Then
    ValidaInterfaz = False
    Exit Function
End If
If CDate(txtIngChqFechaVal) < CDate(txtIngChqFechaReg) Then
    MsgBox "Fecha de Valorizacion no puede ser menor a la de registro", vbInformation, "Aviso"
    txtIngChqFechaVal.SetFocus
    ValidaInterfaz = False
    Exit Function
End If
If CDate(txtIngChqFechaReg) > CDate(txtIngChqFechaVal) Then
    MsgBox "Fecha de Registro no puede ser mayor a la de valorización", vbInformation, "Aviso"
    txtIngChqFechaReg.SetFocus
    ValidaInterfaz = False
    Exit Function
End If

If CDate(txtIngChqFechaReg) > CDate(txtIngChqFechaVal) Then
    MsgBox "Fecha de Registro no puede ser mayor a la de valorización", vbInformation, "Aviso"
    txtIngChqFechaReg.SetFocus
    ValidaInterfaz = False
    Exit Function
End If
If lbNegocio Then
    If DateDiff("d", CDate(txtIngChqFechaReg), CDate(txtIngChqFechaVal)) < lnDiasValoriza Then
        MsgBox "Fecha no válida. Día(s) mínimo(s) de Valorización : [" & lnDiasValoriza & "] ", vbInformation, "Aviso"
        ValidaInterfaz = False
        txtIngChqFechaVal.SetFocus
        Exit Function
    End If
Else
    If txtBuscarProd.Text = "" Or lblProdDesc = "" Then
        MsgBox "Producto a que pertenece el documento no ingresado", vbInformation, "Aviso"
        ValidaInterfaz = False
        If txtBuscarProd.Enabled Then txtBuscarProd.SetFocus
        
        Exit Function
    End If
    If txtBuscarAreaAgencia = "" Or lblAreaAgeDesc = "" Then
        MsgBox "Area/Agencia no ingresada ", vbInformation, "Aviso"
        ValidaInterfaz = False
        If txtBuscarAreaAgencia.Enabled Then txtBuscarAreaAgencia.SetFocus
        Exit Function
    End If
    If txtBuscarCtaHaber.Text = "" Or lblCtaHaber = "" Then
        MsgBox "Cuenta de Haber no Ingresado", vbInformation, "Aviso"
        ValidaInterfaz = False
        If txtBuscarCtaHaber.Enabled Then txtBuscarCtaHaber.SetFocus
        Exit Function
    End If
End If
If Len(Trim(cboMoneda)) = 0 Then
    MsgBox "Moneda de Documento no Seleccionada", vbInformation, "Aviso"
    ValidaInterfaz = False
    cboMoneda.SetFocus
    Exit Function
End If

If Len(Trim(Me.txtMovDesc)) = 0 Then
    MsgBox "Glosa o Descripcion de operación no ingeresada", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtMovDesc.SetFocus
    Exit Function
End If
If Val(txtMonto) = 0 Then
    MsgBox "Monto de Registro no válido", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtMonto.SetFocus
    Exit Function
End If
Set oDoc = Nothing
End Function

Private Sub txtBuscarAreaAgencia_EmiteDatos()
lblAreaAgeDesc = txtBuscarAreaAgencia.psDescripcion
If txtBuscarCtaHaber.Visible And txtBuscarCtaHaber.Enabled Then
    txtBuscarCtaHaber.SetFocus
End If
End Sub

Private Sub txtBuscarAreaAgencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtBuscarCtaHaber.Enabled Then
        txtBuscarCtaHaber.SetFocus
    Else
        txtMovDesc.SetFocus
    End If
End If
End Sub

Private Sub txtBuscarCtaHaber_EmiteDatos()
lblCtaHaber = oContFunct.EmiteCtaContDesc(txtBuscarCtaHaber)
If txtBuscarCtaHaber.Text <> "" Then
    AsignaCtaObj txtBuscarCtaHaber.Text
    If txtMovDesc.Visible And txtMovDesc.Enabled Then
        txtMovDesc.SetFocus
    End If
End If
End Sub

Private Sub txtBuscarCtaHaber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMovDesc.SetFocus
End If
End Sub
Private Sub txtBuscarProd_EmiteDatos()
lblProdDesc = txtBuscarProd.psDescripcion
If txtBuscarAreaAgencia.Visible And txtBuscarAreaAgencia.Enabled Then
    txtBuscarAreaAgencia.SetFocus
End If
End Sub

Private Sub txtBuscarProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtBuscarAreaAgencia.Enabled Then
        txtBuscarAreaAgencia.SetFocus
    End If
End If
End Sub

Private Sub txtIngChqBuscaIF_EmiteDatos()
lblIngChqDescIF = txtIngChqBuscaIF.psDescripcion
If txtIngChqBuscaIF.psDescripcion <> "" Then
    txtIngChqCtaIF.SetFocus
End If
End Sub
Private Sub txtIngChqCtaIF_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    txtIngChqNumCheque.SetFocus
End If
End Sub
Private Sub cboIngChqPlaza_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtIngChqFechaReg.Enabled Then
        txtIngChqFechaReg.SetFocus
    Else
        txtIngChqFechaVal.SetFocus
    End If
End If
End Sub

Private Sub txtIngChqFechaReg_GotFocus()
fEnfoque txtIngChqFechaReg
End Sub

Private Sub txtIngChqFechaReg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtIngChqFechaVal.SetFocus
End If
End Sub
Private Sub txtIngChqFechaReg_Validate(Cancel As Boolean)
If ValFecha(txtIngChqFechaReg) Then
    If lbNegocio Then
        txtIngChqFechaVal = DateAdd("d", lnDiasValoriza, CDate(txtIngChqFechaReg))
    End If
End If
End Sub

Private Sub txtIngChqFechaVal_GotFocus()
fEnfoque txtIngChqFechaVal
End Sub

Private Sub txtIngChqFechaVal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtBuscarProd.Enabled Then
        txtBuscarProd.SetFocus
    Else
        txtMovDesc.SetFocus
    End If
End If
End Sub
Private Sub txtIngChqNumCheque_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    If chkConfirmar.Visible Then
        chkConfirmar.SetFocus
    Else
        cboIngChqPlaza.SetFocus
    End If
End If
End Sub
Public Property Get OK() As Boolean
OK = lbOk
End Property
Public Property Let OK(ByVal vNewValue As Boolean)
lbOk = vNewValue
End Property
Public Property Get PersCodIF() As String
PersCodIF = lsPersCodIf
End Property
Public Property Let PersCodIF(ByVal vNewValue As String)
lsPersCodIf = vNewValue
End Property
Public Property Get NroCtaIf() As String
NroCtaIf = lsNroCtaIf
End Property
Public Property Let NroCtaIf(ByVal vNewValue As String)
lsNroCtaIf = vNewValue
End Property
Public Property Get NroChq() As String
NroChq = lsNroChq
End Property
Public Property Let NroChq(ByVal vNewValue As String)
lsNroChq = vNewValue
End Property
Public Property Get PlazaChq() As String
PlazaChq = lnPlazaChq
End Property
Public Property Let PlazaChq(ByVal vNewValue As String)
lnPlazaChq = vNewValue
End Property
Public Property Get FechaRegChq() As Date
FechaRegChq = ldFechaRegChq
End Property
Public Property Let FechaRegChq(ByVal vNewValue As Date)
ldFechaRegChq = vNewValue
End Property
Public Property Get FechaValChq() As Date
FechaValChq = ldFechaValChq
End Property
Public Property Let FechaValChq(ByVal vNewValue As Date)
FechaValChq = vNewValue
End Property
Public Property Get ConfCheque() As String
ConfCheque = lsConfCheque
End Property
Public Property Let ConfCheque(ByVal vNewValue As String)
lsConfCheque = vNewValue
End Property
Public Property Get CtaContChq() As String
CtaContChq = lsCtaContChq
End Property
Public Property Let CtaContChq(ByVal vNewValue As String)
lsCtaContChq = vNewValue
End Property
Public Property Get Glosa() As String
Glosa = lsGlosa
End Property
Public Property Let Glosa(ByVal vNewValue As String)
lsGlosa = vNewValue
End Property
Public Property Get NombreIF() As String
NombreIF = lsNombreIF
End Property
Public Property Let NombreIF(ByVal vNewValue As String)
NombreIF = vNewValue
End Property

Private Sub txtmonto_GotFocus()
fEnfoque txtMonto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 15, 2)
If KeyAscii = 13 Then
    txtMonto = Format(txtMonto, gsFormatoNumeroView)
    cmdAceptar.SetFocus
End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtMonto.SetFocus
End If
End Sub
Private Sub EmiteDatos()
lsPersCodIf = txtIngChqBuscaIF
lsNroCtaIf = txtIngChqCtaIF
lsNroChq = txtIngChqNumCheque
lnPlazaChq = cboIngChqPlaza.ListIndex
ldFechaRegChq = CDate(txtIngChqFechaReg)
ldFechaValChq = CDate(txtIngChqFechaVal)
lsConfCheque = chkConfirmar.value
lsGlosa = Trim(txtMovDesc)
lsNombreIF = lblIngChqDescIF
lnImporte = CDbl(txtMonto)
Set rsObj = fgObjMotivo.GetRsNew
lsProductoCod = Trim(txtBuscarProd)
lsAreaAgeCod = Trim(txtBuscarAreaAgencia)
lsCtaMotivo = Trim(txtBuscarCtaHaber)
lnMoneda = Val(Right(cboMoneda, 2))
End Sub
Private Sub AsignaDatos()
Dim oDocRec As NDocRec
Dim rs As ADODB.Recordset
Set oDocRec = New NDocRec
If lsPersCodIf = "" Or lsNroChq = "" Then Exit Sub
Set rs = oDocRec.GetDatosCheques(lsNroChq, lsPersCodIf)
If Not rs.EOF And Not rs.BOF Then
    txtIngChqBuscaIF = lsPersCodIf
    lblIngChqDescIF = Trim(rs!IFNOMBRE)
    txtIngChqCtaIF = rs!cIFCta
    txtIngChqNumCheque = lsNroChq
    cboIngChqPlaza.ListIndex = Val(rs!bPlaza)
    txtIngChqFechaReg = Mid(rs!cMovNro, 7, 2) & "/" & Mid(rs!cMovNro, 5, 2) & "/" & Left(rs!cMovNro, 4)
    txtIngChqFechaVal = rs!dValorizaRef
    chkConfirmar.value = rs!nConfCaja
    txtMovDesc = rs!cMovDesc
    lblEstado = rs!cEstado
    txtMonto = Format(rs!nMonto, "#,#0.00")
End If
rs.Close
Set rs = Nothing

End Sub
Private Sub AsignaCtaObj(ByVal psCtaContCod As String)
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim lsRaiz As String
Dim oDescObj As ClassDescObjeto
Dim UP As UPersona
Dim lsFiltro As String
Dim oRHAreas As DActualizaDatosArea
Dim oCtaCont As DCtaCont
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo

Set oEfect = New Defectivo
Set oCtaIf = New NCajaCtaIF
Set oRHAreas = New DActualizaDatosArea
Set oDescObj = New ClassDescObjeto
Set oCtaCont = New DCtaCont
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset

Me.fgObjMotivo.Clear
Me.fgObjMotivo.FormaCabecera
Me.fgObjMotivo.Rows = 2
Set rs1 = oCtaCont.CargaCtaObj(psCtaContCod, , True)
If Not rs1.EOF And Not rs1.BOF Then
    Do While Not rs1.EOF
        lsRaiz = ""
        lsFiltro = ""
        Select Case Val(rs1!cObjetoCod)
            Case ObjCMACAgencias
                Set rs = oRHAreas.GetAgencias(rs1!cCtaObjFiltro)
            Case ObjCMACAgenciaArea
                lsRaiz = "Unidades Organizacionales"
                Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
            Case ObjCMACArea
                Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
            Case ObjEntidadesFinancieras
                lsRaiz = "Cuentas de Entidades Financieras"
                'Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, psCtaContCod)
                Set rs = oCtaIf.CargaCtasIF(Mid(psCtaContCod, 3, 1), rs1!cCtaObjFiltro)
            Case ObjDescomEfectivo
                lsRaiz = "Denominación"
                Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
            Case ObjPersona
                Set rs = Nothing
            Case Else
                Set rs = GetObjetos(Val(rs1!cObjetoCod))
        End Select
        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then
                If Not rs.EOF And Not rs.BOF Then
                    If rs.RecordCount > 1 Then
                        oDescObj.Show rs, "", lsRaiz
                        If oDescObj.lbOk Then
                            lsFiltro = oContFunct.GetFiltroObjetos(Val(rs1!cObjetoCod), psCtaContCod, oDescObj.gsSelecCod, False)
                            fgObjMotivo.AdicionaFila
                            fgObjMotivo.TextMatrix(fgObjMotivo.Row, 1) = oDescObj.gsSelecCod
                            fgObjMotivo.TextMatrix(fgObjMotivo.Row, 2) = oDescObj.gsSelecDesc
                            fgObjMotivo.TextMatrix(fgObjMotivo.Row, 3) = lsFiltro
                            fgObjMotivo.TextMatrix(fgObjMotivo.Row, 4) = rs1!cObjetoCod
                        Else
                            txtBuscarCtaHaber = ""
                            lblCtaHaber = ""
                            Exit Do
                        End If
                    Else
                        fgObjMotivo.AdicionaFila
                        fgObjMotivo.TextMatrix(fgObjMotivo.Row, 1) = rs1!cObjetoCod
                        fgObjMotivo.TextMatrix(fgObjMotivo.Row, 2) = rs1!cObjetoDesc
                        fgObjMotivo.TextMatrix(fgObjMotivo.Row, 3) = lsFiltro
                        fgObjMotivo.TextMatrix(fgObjMotivo.Row, 4) = rs1!cObjetoCod
                    End If
                End If
            End If
        Else
            If Val(rs1!cObjetoCod) = ObjPersona Then
                Set UP = frmBuscaPersona.Inicio
                If Not UP Is Nothing Then
                    fgObjMotivo.AdicionaFila
                    fgObjMotivo.TextMatrix(fgObjMotivo.Row, 1) = UP.sPersCod
                    fgObjMotivo.TextMatrix(fgObjMotivo.Row, 2) = UP.sPersNombre
                    fgObjMotivo.TextMatrix(fgObjMotivo.Row, 3) = ""
                    fgObjMotivo.TextMatrix(fgObjMotivo.Row, 4) = rs1!cObjetoCod
                End If
            End If
        End If
        rs1.MoveNext
    Loop
End If
rs1.Close
Set rs1 = Nothing
Set oDescObj = Nothing
Set UP = Nothing
Set oCtaCont = Nothing
Set oCtaIf = Nothing
Set oEfect = Nothing
End Sub
Public Property Get Importe() As Double
Importe = lnImporte
End Property
Public Property Let Importe(ByVal vNewValue As Double)
lnImporte = vNewValue
End Property
Public Property Get rsObjMotivo() As ADODB.Recordset
Set rsObjMotivo = rsObj
End Property
Public Property Set rsObjMotivo(ByVal vNewValue As ADODB.Recordset)
Set rsObj = vNewValue
End Property
Public Property Get ProductoCod() As String
ProductoCod = lsProductoCod
End Property
Public Property Let ProductoCod(ByVal vNewValue As String)
lsProductoCod = vNewValue
End Property
Public Property Get AreaAgeCod() As String
AreaAgeCod = lsAreaAgeCod
End Property
Public Property Let AreaAgeCod(ByVal vNewValue As String)
lsAreaAgeCod = vNewValue
End Property
Public Property Get CtaMotivo() As String
CtaMotivo = lsCtaMotivo
End Property
Public Property Let CtaMotivo(ByVal vNewValue As String)
lsCtaMotivo = vNewValue
End Property
Public Property Get Moneda() As Moneda
Moneda = lnMoneda
End Property
Public Property Let Moneda(ByVal vNewValue As Moneda)
lnMoneda = vNewValue
End Property
