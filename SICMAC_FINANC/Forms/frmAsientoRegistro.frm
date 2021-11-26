VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAsientoRegistro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Asientos Contables"
   ClientHeight    =   7650
   ClientLeft      =   -1995
   ClientTop       =   435
   ClientWidth     =   9300
   Icon            =   "frmAsientoRegistro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDistribucion 
      Caption         =   "Inserta Cta. Distribuida"
      Height          =   375
      Left            =   7080
      TabIndex        =   35
      Top             =   240
      Width           =   2055
   End
   Begin VB.CheckBox chkDoc 
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
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   750
      Width           =   1365
   End
   Begin Sicmact.FlexEdit fgDoc 
      Height          =   735
      Left            =   90
      TabIndex        =   3
      Top             =   870
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   1296
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Tipo-Descripción-Numero-Fecha-nMovNro-cMovNro"
      EncabezadosAnchos=   "400-1200-3530-2100-1500-0-0"
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
      ColumnasAEditar =   "X-1-X-3-4-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-1-0-0-2-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-L-L-C-L-L"
      FormatosEdit    =   "0-3-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   2715
      Left            =   90
      TabIndex        =   7
      Top             =   2700
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   4789
      Cols0           =   8
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "#-Código-Descripción-ItemCtaCont-DEBE S/.-HABER S/.-DEBE $-HABER $"
      EncabezadosAnchos=   "385-1700-3800-0-1400-1400-1400-1400"
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
      ColumnasAEditar =   "X-1-X-X-4-5-6-7"
      TextStyleFixed  =   3
      ListaControles  =   "0-1-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-R-R-R-R"
      FormatosEdit    =   "0-0-0-0-2-2-2-2"
      CantEntero      =   15
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   390
      RowHeight0      =   285
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdCancelarExt 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7980
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   14
      Top             =   7230
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   360
      Left            =   90
      TabIndex        =   11
      Top             =   7230
      Width           =   1095
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "E&liminar"
      Height          =   360
      Left            =   1230
      TabIndex        =   12
      Top             =   7230
      Width           =   1095
   End
   Begin VB.Frame fraDato 
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
      Left            =   90
      TabIndex        =   19
      Top             =   1560
      Width           =   9075
      Begin VB.TextBox txtPersNombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2820
         TabIndex        =   5
         Top             =   240
         Width           =   5985
      End
      Begin Sicmact.TxtBuscar txtPersCod 
         Height          =   330
         Left            =   870
         TabIndex        =   4
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
      End
      Begin VB.TextBox txtMovDesc 
         Height          =   345
         Left            =   870
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   630
         Width           =   7935
      End
      Begin VB.Label Label5 
         Caption         =   "Persona"
         Height          =   255
         Left            =   210
         TabIndex        =   27
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Glosa"
         Height          =   255
         Left            =   210
         TabIndex        =   26
         Top             =   660
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operación"
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
      Height          =   645
      Left            =   90
      TabIndex        =   18
      Top             =   60
      Width           =   6855
      Begin VB.ComboBox cboOperacion 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   4215
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   5520
         TabIndex        =   1
         Top             =   210
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   4950
         TabIndex        =   22
         Top             =   285
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   150
         TabIndex        =   21
         Top             =   270
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7980
      TabIndex        =   16
      Top             =   7230
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   6840
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   15
      Top             =   7230
      Width           =   1095
   End
   Begin Sicmact.FlexEdit fgObj 
      Height          =   945
      Left            =   90
      TabIndex        =   8
      Top             =   6180
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   1667
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "#-Orden-Código-Descripción-cObjetoCod-ItemCtaCont"
      EncabezadosAnchos=   "385-700-1800-5850-0-0"
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
      ColumnasAEditar =   "X-X-2-X-X-X"
      TextStyleFixed  =   1
      ListaControles  =   "0-0-1-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-L-L-C"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   390
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame fraMovDesc 
      Caption         =   "Datos del Extorno"
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
      Height          =   1035
      Left            =   90
      TabIndex        =   23
      Top             =   6060
      Visible         =   0   'False
      Width           =   9060
      Begin VB.TextBox txtMovDescExt 
         Height          =   585
         Left            =   3000
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   300
         Width           =   5835
      End
      Begin MSMask.MaskEdBox txtFechaExt 
         Height          =   315
         Left            =   780
         TabIndex        =   9
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Caption         =   "Motivo"
         Height          =   240
         Left            =   2340
         TabIndex        =   25
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   180
         TabIndex        =   24
         Top             =   330
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdAceptarExt 
      Caption         =   "&Extornar"
      Height          =   360
      Left            =   6840
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   13
      Top             =   7230
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label txtSaldoD 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7560
      TabIndex        =   34
      Top             =   5760
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label txtSaldo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4710
      TabIndex        =   33
      Top             =   5760
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label txtHaberD 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   7560
      TabIndex        =   32
      Top             =   5490
      Width           =   1350
   End
   Begin VB.Label txtDebeD 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6180
      TabIndex        =   31
      Top             =   5490
      Width           =   1350
   End
   Begin VB.Label txtHaber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.00"
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
      Height          =   255
      Left            =   4710
      TabIndex        =   30
      Top             =   5490
      Width           =   1350
   End
   Begin VB.Label txtDebe 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.00"
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
      Height          =   255
      Left            =   3270
      TabIndex        =   29
      Top             =   5490
      Width           =   1350
   End
   Begin VB.Label lblObj 
      AutoSize        =   -1  'True
      Caption         =   "Objetos"
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
      Left            =   60
      TabIndex        =   20
      Top             =   5910
      Width           =   660
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TOTALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2280
      TabIndex        =   17
      Top             =   5490
      Width           =   945
   End
   Begin VB.Label lblTotalC 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2130
      TabIndex        =   28
      Top             =   5430
      Width           =   7005
   End
End
Attribute VB_Name = "frmAsientoRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim rsCta As ADODB.Recordset
Dim rs    As ADODB.Recordset
Dim sMovNroAnt    As String
Dim nMovNroAnt    As Long
Dim lbExtornar As Boolean
Dim nTipCambio As Currency
Dim nTipCambioC As Currency
Dim nTipCambioV As Currency
Dim lbPorMoneda As Boolean

Dim lbCargaCuentas      As Boolean
Dim lbRegistraPendiente As Boolean
Dim lbRegulaPendiente   As Boolean
Dim lnSaldo             As Currency
Dim OK As Boolean
Dim lsClasePendiente As String
Dim lnNroCta As Integer

'*** PEAC 20100712
Dim sMatrizDatos() As Variant
Dim nTipoMatriz As Integer
Dim nCantAgeSel As Integer
Dim nMontogasto As Currency
Dim sCuentaDistrib As Currency
Dim sCtaContrapDistriAstoManual As String
'*** FIN PEAC

Public sPersona As String
Public sNomPers As String
Dim oContFunct As NContFunciones
Dim oCta        As DCtaCont
Dim lsAgeCodRef As String
Dim rsPend      As ADODB.Recordset

'ARLO20170208****
Dim objPista As COMManejador.Pista
Dim lsPalabra, lsMovAnt, lscod, lsAccion As String
Dim lnNumero, i As Integer
'************

Public Sub Inicio(psMov As String, pnMov As Long, Optional pbExtornar As Boolean = False, Optional pbPorMoneda As Boolean = False, Optional pbCargaCuentas As Boolean = False, Optional pbRegistraPendiente As Boolean = False, Optional pbRegulaPendiente As Boolean = False, Optional psAgeRef As String = "", Optional prsPend As ADODB.Recordset)
sMovNroAnt = psMov
nMovNroAnt = pnMov
lbExtornar = pbExtornar
lbPorMoneda = pbPorMoneda

'Nuevas para Pendientes
lbCargaCuentas = pbCargaCuentas
lbRegistraPendiente = pbRegistraPendiente
lbRegulaPendiente = pbRegulaPendiente
lsAgeCodRef = psAgeRef
Set rsPend = prsPend
Me.Show 1
End Sub

Private Sub EliminaCuenta(sCod As String, nItem As Integer)
fg.EliminaFila fg.row, False
EliminaFgObj nItem
If Len(fg.TextMatrix(1, 1)) > 0 Then
   RefrescaFgObj fg.TextMatrix(fg.row, 0)
   Sumas
Else
   txtDebe = ""
   txtHaber = ""
End If
End Sub

Private Sub Sumas()
   txtDebe = Format(fg.SumaRow(4), gsFormatoNumeroView)
   txtHaber = Format(fg.SumaRow(5), gsFormatoNumeroView)
   If fg.Cols > 6 Then
      txtDebeD = Format(fg.SumaRow(6), gsFormatoNumeroView)
      txtHaberD = Format(fg.SumaRow(7), gsFormatoNumeroView)
      VerSaldo txtDebeD, txtHaberD, txtSaldoD
   End If
   VerSaldo txtDebe, txtHaber, txtSaldo
End Sub

Private Sub VerSaldo(txtD As Label, txtH As Label, txt As Label)
   If txtH <> txtD Then
      txt.Visible = True
      txt = Format(Abs(nVal(txtD.Caption) - nVal(txtH.Caption)), gsFormatoNumeroView)
      If nVal(txtH.Caption) > nVal(txtD.Caption) Then
         txt.Left = txtD.Left
      Else
         txt.Left = txtH.Left
      End If
   Else
      txt.Visible = False
   End If
End Sub

Private Function ValidaDatos() As Boolean
Dim i As Integer
ValidaDatos = False
If cboOperacion.ListIndex = -1 Then
   MsgBox "Falta seleccionar Tipo de Operación...!", vbInformation, "¡Aviso!"
   cboOperacion.SetFocus
   Exit Function
End If
If Len(fg.TextMatrix(1, 1)) = 0 Or (Val(txtDebe) = 0 And Val(txtHaber) = 0 And Val(txtDebeD) = 0 And Val(txtHaberD) = 0) Then
   MsgBox " Asiento vacio...! ", vbInformation, "¡Aviso!"
   fg.SetFocus
   Exit Function
End If

For i = 1 To Me.fg.Rows - 1
'*** Cambio de Código Para Ingreso Manual de Asientos

If gsOpeCod = "701100" Then
    If fg.TextMatrix(i, 1) = "" And (fg.TextMatrix(i, 4) <> "" Or fg.TextMatrix(i, 5) <> "" Or fg.TextMatrix(i, 6) <> "" Or fg.TextMatrix(i, 7) <> "") Then
        MsgBox " Cuenta en Blanco !!!!! ", vbInformation, "¡Aviso!"
        fg.row = i
        fg.col = 1
        fg.SetFocus
        Exit Function
    End If
ElseIf gsOpeCod = "701160" Then
     If fg.TextMatrix(i, 1) = "" And (fg.TextMatrix(i, 4) <> "" Or fg.TextMatrix(i, 5) <> "") Then
        MsgBox " Cuenta en Blanco !!!!! ", vbInformation, "¡Aviso!"
        fg.row = i
        fg.col = 1
        fg.SetFocus
        Exit Function
    End If

End If
Next i

i = 1
Do While i < fgDoc.Rows
    If fgDoc.TextMatrix(i, 0) = "" Then
        i = i + 1
    Else
        If fgDoc.TextMatrix(i, 1) = "" Then
           fgDoc.EliminaFila i
        Else
            If fgDoc.TextMatrix(i, 3) = "" Then
                MsgBox "Falta Número de Documento", vbInformation, "¡Aviso!"
                fgDoc.row = i
                fgDoc.col = 3
                fgDoc.SetFocus
                Exit Function
            End If
            If fgDoc.TextMatrix(i, 4) = "" Then
                MsgBox "Falta Fecha de Documento", vbInformation, "¡Aviso!"
                fgDoc.row = i
                fgDoc.col = 4
                fgDoc.SetFocus
                Exit Function
            End If
            i = i + 1
        End If
    End If
Loop

'Regulariza Pendientes
If lbRegulaPendiente Then
    If Mid(gsOpeCod, 3, 1) = 2 Then
        If lsClasePendiente = "D" And (nVal(fg.TextMatrix(1, 7)) = 0 Or nVal(fg.TextMatrix(1, 6)) <> 0) Then
            MsgBox "Regularización de Pendiente debe ingresarse en el Haber", vbInformation, "¡Aviso!"
            Exit Function
        End If
        If lsClasePendiente = "A" And (nVal(fg.TextMatrix(1, 6)) = 0 Or nVal(fg.TextMatrix(1, 7)) <> 0) Then
            MsgBox "Regularización de Pendiente debe ingresarse en el Debe", vbInformation, "¡Aviso!"
            Exit Function
        End If
    Else
        If lsClasePendiente = "D" And (nVal(fg.TextMatrix(1, 5)) = 0 Or nVal(fg.TextMatrix(1, 4)) <> 0) Then
            MsgBox "Regularización de Pendiente debe ingresarse en el Haber", vbInformation, "¡Aviso!"
            Exit Function
        End If
        If lsClasePendiente = "A" And (nVal(fg.TextMatrix(1, 4)) = 0 Or nVal(fg.TextMatrix(1, 5)) <> 0) Then
            MsgBox "Regularización de Pendiente debe ingresarse en el Debe", vbInformation, "¡Aviso!"
            Exit Function
        End If
    End If
End If

If lbRegistraPendiente Then
    Dim lnMontoPendiente As Currency
    For i = 1 To lnNroCta
        If lsClasePendiente = "A" Then
            If Mid(gsOpeCod, 3, 1) = 1 Then
                lnMontoPendiente = lnMontoPendiente + nVal(fg.TextMatrix(i, 5))
            Else
                lnMontoPendiente = lnMontoPendiente + nVal(fg.TextMatrix(i, 7))
            End If
        End If
        If lsClasePendiente = "D" Then
            If Mid(gsOpeCod, 3, 1) = 1 Then
                lnMontoPendiente = lnMontoPendiente + nVal(fg.TextMatrix(i, 4))
            Else
                lnMontoPendiente = lnMontoPendiente + nVal(fg.TextMatrix(i, 6))
            End If
        End If
    Next
    If lnMontoPendiente = 0 Then
        MsgBox "No se Registro monto para Cuentas de Pendiente", vbInformation, "¡Aviso!"
        Exit Function
    End If
End If

For i = 1 To fgObj.Rows - 1
   If fgObj.TextMatrix(i, 4) <> "" And fgObj.TextMatrix(i, 4) <> "00" And fgObj.TextMatrix(i, 2) = "" Then
      MsgBox "Falta información de Objeto relacionado con Cuenta Contable", vbInformation, "¡Aviso!"
      Exit Function
   End If
Next
If txtMovDesc = "" Then
   MsgBox "Falta ingresar descripción de la Operación", vbInformation, "¡Aviso!"
   txtMovDesc.SetFocus
   Exit Function
End If
If ValidaFecha(txtFecha) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Aviso"
   Exit Function
Else
   If Val(Right(txtFecha, 4)) < Year(gdFecSis) And sMovNroAnt = "" Then
      If MsgBox("Año no corresponde a Ejercicio...¿DESEA CONTINUAR?", vbQuestion + vbYesNo, "!Aviso!") = vbNo Then
         txtFecha.SetFocus
         Exit Function
      End If
   End If
End If
If txtHaber <> txtDebe Then
   MsgBox "Suma de Debe y Haber en Soles no son iguales. Verificar", vbInformation, "¡Aviso!"
   fg.SetFocus
   Exit Function
End If
ValidaDatos = True
End Function

Private Sub cboOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFecha.SetFocus
End If
End Sub

Private Sub chkDoc_Click()
If chkDoc.value = vbChecked Then
    fgDoc.lbEditarFlex = True
    If fgDoc.TextMatrix(1, 1) = "" Then
        fgDoc.AdicionaFila
    End If
Else
    fgDoc.lbEditarFlex = False
End If
End Sub

Private Sub cmdAceptar_Click()
Dim nItem As Integer
Dim nImporte   As Currency
Dim sMovNroModifica As String
Dim clsmov As New DMov
Dim oDocPago As New clsDocPago
Dim lsDocNroVoucher As String
Dim ldFechaVoucher As Date
Dim lsDocumento As String
Dim lsPiePag    As String
On Error GoTo ErrAceptar

If Not ValidaDatos Then
   Exit Sub
End If
If sMovNroAnt = "" Then
   gsMovNro = clsmov.GeneraMovNro(txtFecha, Right(gsCodAge, 2), gsCodUser)
   sMovNroModifica = ""
Else
   gsMovNro = clsmov.GeneraMovNro(txtFecha, , , sMovNroAnt)
   sMovNroModifica = clsmov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
End If
gsOpeCod = Mid(cboOperacion.Text, 1, 6)
gdFecha = txtFecha
gsGlosa = txtMovDesc
gnDocTpo = -1
gnImporte = 0
lsPiePag = "9"
If fgDoc.TextMatrix(1, 1) <> "" Then
   Select Case fgDoc.TextMatrix(1, 1)
    Case TpoDocOrdenPago   'Orden Pago
        gnDocTpo = TpoDocOrdenPago
        gnImporte = IIf(Mid(gsOpeCod, 3, 1) = "1", CCur(txtDebe), CCur(txtDebeD))
        Screen.MousePointer = 11
        gnDocTpo = TpoDocOrdenPago
        oDocPago.InicioOrdenPago fgDoc.TextMatrix(1, 3), True, txtPersCod, gsOpeCod, txtPersNombre, gsOpeCod, txtMovDesc, gnImporte, gdFecSis, lsDocNroVoucher, False ', gsCodAge
        Screen.MousePointer = 0
        If oDocPago.vbOk Then    'Se ingresó dato de Cheque u Orden de Pago
            fgDoc.TextMatrix(1, 4) = oDocPago.vdFechaDoc
            fgDoc.TextMatrix(1, 3) = oDocPago.vsNroDoc
            lsDocNroVoucher = oDocPago.vsNroVoucher
            ldFechaVoucher = oDocPago.vdFechaDoc
            lsDocumento = oDocPago.vsFormaDoc
            gnImporte = oDocPago.vnImporte
            txtMovDesc = oDocPago.vsGlosa
            lsPiePag = "17"
        Else
            Exit Sub
        End If
   End Select
End If

If MsgBox(" ¿ Seguro de grabar Operación ? ", vbOKCancel, "Aviso de Confirmación") = vbCancel Then
   Exit Sub
End If
MousePointer = 11
cmdAceptar.Enabled = False

Dim oMov As New ncontasientos
If lbPorMoneda Then
    oMov.GrabaAsientoContableMoneda gsMovNro, gsOpeCod, txtMovDesc, gnImporte, fgDoc.GetRsNew, fg.GetRsNew, fgObj.GetRsNew, sMovNroModifica, sMovNroAnt, nMovNroAnt, txtPersCod.psCodigoPersona, lbRegistraPendiente, lbRegulaPendiente, lsAgeCodRef, gbBitCentral, rsPend
Else
    oMov.GrabaAsientoContable gsMovNro, gsOpeCod, txtMovDesc, gnImporte, fgDoc.GetRsNew, fg.GetRsNew, fgObj.GetRsNew, sMovNroModifica, sMovNroAnt, nMovNroAnt, txtPersCod.psCodigoPersona, Me.txtPersNombre.Tag
End If
            'ARLO20170208
            Dim lnNumero, i As Integer
            If (lbPorMoneda) Then
            lsPalabra = "Asiento Contable Con Moneda"
            Else: lsPalabra = "Asiento Contable"
            End If
            If (sMovNroAnt) = "" Then
            lsMovAnt = "Agrego"
            lsAccion = "1"
            Else: lsMovAnt = "Modifico"
            lsAccion = "2"
            End If
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaRegistoAsientCont
            lnNumero = fg.Rows - 1
            lscod = ""
            For i = 1 To lnNumero
            lscod = lscod + " , " + fg.TextMatrix(i, 1)
            Next i
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, "Se " & lsMovAnt & " " & lsPalabra & " | Tipo de Operación : " & cboOperacion.Text & " | Asiento Contable : " & Replace(lscod, ",", " ")
            Set objPista = Nothing
            '*******
            
'FECHA DEL SALDO
'If CDate(txtFecha) < gdFecSis Then
'   Dim oFun As New NConstSistemas
'       oFun.ActualizaConstSistemas gConstSistUltActSaldos, GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge), txtFecha
'   Set oFun = Nothing
'End If

' Imprimimos ASIENTO CONTABLE
Me.Enabled = False
MousePointer = 0
If Val(fgDoc.TextMatrix(1, 1)) = TpoDocRecEgreso Then
    Dim lsRec As String
    Dim oImp As NContImprimir
    Set oImp = New NContImprimir
    oImp.Inicio gsNomCmac, gsNomAge, Format(gdFecSis, gsFormatoFechaView)
    lsRec = oImp.ImprimeRecibo(gsMovNro)
    Set oImp = Nothing
    EnviaPrevio lsRec, "IMPRESION DE DOCUMENTO", gnLinPage, False
End If
ImprimeAsientoContable gsMovNro, lsDocNroVoucher, gnDocTpo, lsDocumento, , , txtMovDesc, txtPersCod, gnImporte, , , , 1, , lsPiePag
Me.Enabled = True
cmdAceptar.Enabled = True
glAceptar = True
OK = True
If sMovNroAnt <> "" Or lbRegulaPendiente Then
   Unload Me
   Exit Sub
End If
If MsgBox(" ¿ Desea registrar otro Asiento ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   fg.Clear
   fg.Rows = 2
   fg.FormaCabecera
   fg.FormateaColumnas
   fg.TextMatrix(1, 0) = "1"
   fgObj.Clear
   fgObj.Rows = 2
   fgObj.FormaCabecera
   fgObj.FormateaColumnas
   If lbPorMoneda Then
      If gsSimbolo = gcMN Then
         fg.Cols = 6
      Else
         fg.ColWidth(4) = 0
         fg.ColWidth(5) = 0
      End If
   End If
   fgDoc.Clear
   fgDoc.Rows = 2
   fgDoc.FormaCabecera
   fgDoc.FormateaColumnas
   chkDoc.value = vbUnchecked
   txtPersCod = ""
   txtPersCod.psCodigoPersona = ""
   txtMovDesc = ""
   txtPersNombre = ""
   Sumas
   fg.SetFocus
Else
   Unload Me
End If
Exit Sub
ErrAceptar:
   MousePointer = 0
   MsgBox TextErr(Err.Description), vbInformation, "Error de Actualizacion"
   cmdAceptar.Enabled = True
End Sub

Private Sub cmdAceptarExt_Click()
Dim nPos       As Variant
Dim sMovNroModifica As String
Dim sMovCambio As String
Dim sImpre     As String

If txtMovDescExt = "" Then
   MsgBox "Falta indicar motivo de Extorno...!", vbInformation, "¡Aviso!"
   txtMovDescExt.SetFocus
   Exit Sub
End If
If MsgBox(" ¿ Seguro que desea Extornar Movimiento ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbYes Then
   Dim oFun As New NContFunciones
   sMovNroModifica = oFun.GeneraMovNro(CDate(txtFechaExt), gsCodAge, gsCodUser)
   If Not oFun.PermiteModificarAsiento(sMovNroModifica, False) Then
      MsgBox "Fecha de Extorno corresponde a un mes ya Cerrado", vbInformation, "¡Aviso!"
      Exit Sub
   End If
   Set oFun = Nothing
   
   Dim oMov As New DMov
   oMov.BeginTrans
   sMovCambio = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
   oMov.ExtornaMovimiento sMovNroModifica, nMovNroAnt, Left(cboOperacion, 6), txtMovDescExt, , sMovCambio
   oMov.CommitTrans
   glAceptar = True
   Set oMov = Nothing
   
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            lscod = ""
            lnNumero = fg.Rows - 1
            For i = 1 To lnNumero
            lscod = lscod + " , " + fg.TextMatrix(i, 1)
            Next i
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se Extorno el Asiento Contable " & lscod & "de la Tipo de Operación : " & cboOperacion.Text & " por Motivo " & txtMovDescExt
            Set objPista = Nothing
            '*******
   Dim oImpre As NContImprimir
   Set oImpre = New NContImprimir
   sImpre = oImpre.ImprimeAsientoContable(sMovNroModifica, gnLinPage, gnColPage, "ASIENTO CONTABLE DE EXTORNO")
   Set oImpre = Nothing
   EnviaPrevio sImpre, "ASIENTO CONTABLE DE EXTORNO", gnLinPage, False
   Unload Me
End If
End Sub

Private Sub cmdCancelarExt_Click()
Unload Me
glAceptar = False
End Sub

Private Sub cmdCerrar_Click()
Unload Me
glAceptar = False
End Sub

'*** PEAC 20100906
Private Sub cmdDistribucion_Click()
         
    Dim oMov As DMov
    Dim oDescObj As New ClassDescObjeto
    Dim rsObj As ADODB.Recordset
    Dim oCta As DCtaCont
    Dim lcCuenta As String, lcCtaDesc As String, lcCtaDistrib As String
    Dim rsAgesDistrib As ADODB.Recordset
    Dim lsSql As String, lnItemDistri As Integer
    Dim m_lbUltimaInstancia As Boolean
    Set oCta = New DCtaCont
    
    Dim lcCtaContraDesc As String
    Dim lnMontoTot As Currency
    Dim lnMonto As Currency
    Dim lnMontoDif As Currency
    Dim nContx As Integer
    
    Set rsObj = oCta.CargaCtaCont("cCtaContCod LIKE '__[^0]%' and nCtaEstado=1", , adLockReadOnly, True)
    Set oCta = Nothing

    oDescObj.lbUltNivel = m_lbUltimaInstancia
    oDescObj.ShowGrid rsObj, "Cuentas"
    rsObj.Close

    If oDescObj.lbOk Then
        'MsgBox oDescObj.gsSelecCod + " - " + oDescObj.gsSelecDesc
        lcCuenta = Trim(oDescObj.gsSelecCod)
        lcCtaDesc = Trim(oDescObj.gsSelecDesc)
    Else
        lcCuenta = ""
        lcCtaDesc = ""
        Exit Sub
    End If

    If ValidaCtaAge(lcCuenta) Then
        If Right(lcCuenta, 2) = "01" Then
            Call frmAgenciaPorcentajeGastosProvision.Inicio(sMatrizDatos, nTipoMatriz, nCantAgeSel, nMontogasto, sCtaContrapDistriAstoManual, True)
            If nCantAgeSel > 0 Then
                'lbPorMoneda
                
                    Dim oCont As DCtaCont
                    Set oCont = New DCtaCont

                    Set oMov = New DMov
                    Set rsAgesDistrib = New ADODB.Recordset
                    lsSql = " exec stp_sel_AgenciaPorcentajeGastos "
                    Set rsAgesDistrib = oMov.CargaRecordSet(lsSql)
                    Set oMov = Nothing
                    
                    lnItemDistri = 1
                    lnMontoTot = 0
                    
                    '*** nueva forma de distribuir  - PEAC
                    For nContx = 1 To nCantAgeSel
                        lcCtaDistrib = Left(lcCuenta, Len(lcCuenta) - 2) + sMatrizDatos(1, lnItemDistri) ''rsAgesDistrib!cAgecod
                            If oCont.ExisteCuenta(lcCtaDistrib) Then
                                fg.AdicionaFila
                                fg.TextMatrix(fg.row, 1) = lcCtaDistrib
                                fg.TextMatrix(fg.row, 2) = sMatrizDatos(2, lnItemDistri) ''rsAgesDistrib!cAgeDescripcion
                                fg.TextMatrix(fg.row, 3) = fg.Rows - 1
                                
                                If gsOpeCod = "701160" Or gsOpeCod = "701100" Or gsOpeCod = "702100" Then
                                    'lnMonto = Round((nMontogasto * rsAgesDistrib!nAgePorcentaje) / 100, 2)
                                    lnMonto = Round((nMontogasto * sMatrizDatos(3, lnItemDistri)) / 100, 2)
                                    fg.TextMatrix(fg.row, 4) = lnMonto
                                ElseIf gsOpeCod = "702160" Then
                                    lnMonto = Round((nMontogasto * sMatrizDatos(3, lnItemDistri)) / 100, 2)
                                    fg.TextMatrix(fg.row, 6) = lnMonto
                                End If
                                
                                lnMontoTot = lnMontoTot + lnMonto
                                
                                lnItemDistri = lnItemDistri + 1
'                                If lnItemDistri > nCantAgeSel Then
'                                    Exit Do
'                                End If
                            End If
                    Next nContx
                     
'                    Do While Not rsAgesDistrib.EOF
'                        If rsAgesDistrib!cAgecod = sMatrizDatos(1, lnItemDistri) Then
'                            lcCtaDistrib = Left(lcCuenta, Len(lcCuenta) - 2) + rsAgesDistrib!cAgecod
'                            If oCont.ExisteCuenta(lcCtaDistrib) Then
'                                fg.AdicionaFila
'                                fg.TextMatrix(fg.Row, 1) = lcCtaDistrib
'                                fg.TextMatrix(fg.Row, 2) = rsAgesDistrib!cAgeDescripcion
'                                fg.TextMatrix(fg.Row, 3) = fg.Rows - 1
'
'                                If gsOpeCod = "701160" Or gsOpeCod = "701100" Or gsOpeCod = "702100" Then
'                                    lnMonto = Round((nMontogasto * rsAgesDistrib!nAgePorcentaje) / 100, 2)
'                                    fg.TextMatrix(fg.Row, 4) = lnMonto
'                                ElseIf gsOpeCod = "702160" Then
'                                    lnMonto = Round((nMontogasto * rsAgesDistrib!nAgePorcentaje) / 100, 2)
'                                    fg.TextMatrix(fg.Row, 6) = lnMonto
'                                End If
'
'                                lnMontoTot = lnMontoTot + lnMonto
'
'                                lnItemDistri = lnItemDistri + 1
'                                If lnItemDistri > nCantAgeSel Then
'                                    Exit Do
'                                End If
'                            End If
'                        End If
'                        rsAgesDistrib.MoveNext
'                    Loop
                    
                    If nMontogasto > lnMontoTot Then
                        lnMontoDif = nMontogasto - lnMontoTot
                        If gsOpeCod = "701160" Or gsOpeCod = "701100" Or gsOpeCod = "702100" Then
                            fg.TextMatrix(fg.row, 4) = lnMonto + lnMontoDif
                        ElseIf gsOpeCod = "702160" Then
                            fg.TextMatrix(fg.row, 6) = lnMonto + lnMontoDif
                        End If
                    ElseIf lnMontoTot > nMontogasto Then
                        lnMontoDif = lnMontoTot - nMontogasto
                        If gsOpeCod = "701160" Or gsOpeCod = "701100" Or gsOpeCod = "702100" Then
                            fg.TextMatrix(fg.row, 4) = lnMonto - lnMontoDif
                        ElseIf gsOpeCod = "702160" Then
                            fg.TextMatrix(fg.row, 6) = lnMonto - lnMontoDif
                        End If
                    End If

                    If Len(sCtaContrapDistriAstoManual) > 0 Then
                        
                        Set oCta = New DCtaCont
                        lcCtaContraDesc = oCta.GetCtaContDesc(sCtaContrapDistriAstoManual)
                        
                        fg.AdicionaFila
                        fg.TextMatrix(fg.row, 1) = sCtaContrapDistriAstoManual
                        fg.TextMatrix(fg.row, 2) = lcCtaContraDesc
                        fg.TextMatrix(fg.row, 3) = fg.Rows - 1
                    
                        If gsOpeCod = "701160" Or gsOpeCod = "701100" Or gsOpeCod = "702100" Then
                            fg.TextMatrix(fg.row, 5) = nMontogasto
                        ElseIf gsOpeCod = "702160" Then
                            fg.TextMatrix(fg.row, 7) = nMontogasto
                        End If
                    
                        Set oCta = Nothing
                    
                    End If
                    
                    RSClose rsAgesDistrib
                    Set rsAgesDistrib = Nothing
            Else
                Exit Sub
            End If
        Else
            MsgBox "Solo se puede distribuir de la Agencia Principal.", vbOKOnly, "Atención"
        End If
    Else
        MsgBox "Cuenta no se puede Distribuir porque no tiene Agencias.", vbOKOnly, "Atención"
    End If
    Sumas
End Sub

Private Sub cmdEliminar_Click()
If fg.TextMatrix(fg.row, 0) <> "" Then
   EliminaCuenta fg.TextMatrix(fg.row, 1), fg.TextMatrix(fg.row, 0)
   If fg.TextMatrix(1, 0) = "" Then
      fg.TextMatrix(1, 0) = "1"
   End If
   If fg.Enabled Then
      fg.SetFocus
   End If
End If
End Sub

Private Sub cmdNuevo_Click()
fg.AdicionaFila , Val(fg.TextMatrix(fg.Rows - 1, 0)) + 1
fg.SetFocus
End Sub

Private Sub Command1_Click()

End Sub

Private Sub fg_OnCellChange(pnRow As Long, pnCol As Long)
Dim nMonto As Currency
Dim nImporte As Currency
If Not lbPorMoneda Then
    If pnCol > 3 And pnCol < 6 And Mid(fg.TextMatrix(pnRow, 1), 3, 1) = "2" Then
       If nTipCambio = 0 Then
          MsgBox "No se definio Tipo de Cambio", vbInformation, "¡Aviso!"
          Exit Sub
       End If
       nImporte = Val(Format(fg.TextMatrix(pnRow, pnCol), gsFormatoNumeroDato))
       Select Case Mid(fg.TextMatrix(pnRow, 1), 1, 1)
              Case "1", "2", "3", "7", "8": nImporte = Round(nImporte / nTipCambio, 2) 'Activo Pasivo
              Case "4": nImporte = Round(nImporte / nTipCambioV, 2)     'Gasto
              Case "5": nImporte = Round(nImporte / nTipCambioC, 2)     'Ingreso
       End Select
       Select Case Mid(fg.TextMatrix(pnRow, 1), 1, 2)
              Case "63", "65": nImporte = Round(nImporte / nTipCambioV, 2)     'Gasto
              Case "62", "64": nImporte = Round(nImporte / nTipCambioC, 2)    'Ingreso
       End Select
       fg.TextMatrix(pnRow, IIf(pnCol = 4, 6, 7)) = Format(nImporte, gsFormatoNumeroView)
    End If
End If
    If pnCol > 1 And fg.TextMatrix(pnRow, 1) = "" Then
        MsgBox "Debe ingresar primero Cuenta Contable", vbInformation, "¡Aviso!"
        fg.TextMatrix(pnRow, pnCol) = ""
    End If
Sumas
fg.SetFocus
End Sub

Private Sub fg_OnRowAdd(pnRow As Long)
   RefrescaFgObj fg.TextMatrix(pnRow, 0)
End Sub

Private Sub fg_OnRowChange(pnRow As Long, pnCol As Long)
If Len(fg.TextMatrix(1, 1)) > 0 Then
  ' pnRow = fgObj.TextMatrix(1, 0)
   RefrescaFgObj fg.TextMatrix(fg.row, 0)
End If
End Sub

Private Sub fgDoc_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim oFun As New NContFunciones
If psDataCod <> "" Then
   Select Case psDataCod
       Case TpoDocRecEgreso
            fgDoc.TextMatrix(pnRow, pnCol + 2) = oFun.GeneraDocNro(Val(psDataCod), gsCodUser)
       Case TpoDocOrdenPago
            fgDoc.TextMatrix(pnRow, pnCol + 2) = oFun.GeneraDocNro(TpoDocOrdenPago, , Mid(gsOpeCod, 3, 1), Right(gsCodAge, 2))
   End Select
    fgDoc.TextMatrix(pnRow, pnCol + 3) = txtFecha
End If
Set oFun = Nothing
End Sub

Private Sub Form_Load()
Dim n As Integer
Set oCta = New DCtaCont

nTipCambio = gnTipCambio
nTipCambioV = gnTipCambioV
nTipCambioC = gnTipCambioC

'*** PEAC 20100713 - solo muestra el botn de distribucion en soles
If gsOpeCod = "701100" Or gsOpeCod = "701160" Or gsOpeCod = "702100" Or gsOpeCod = "702160" Then
    Me.cmdDistribucion.Visible = True
Else
    Me.cmdDistribucion.Visible = False
End If
'*** FIN PEAC

If Not lbExtornar Then
'ALPA 20090930***************************************************
   Set rs = oCta.CargaCtaCont("cCtaContCod LIKE '__[^0]%' and nCtaEstado=1", , adLockReadOnly, True)
'****************************************************************
   fg.TipoBusqueda = BuscaGrid
   fg.lbUltimaInstancia = True
   fg.AutoAdd = True
   fg.rsTextBuscar = rs
   txtPersCod.TipoBusPers = BusPersDocumentoRuc
   txtPersCod.TipoBusqueda = BuscaPersona
   txtPersCod.EditFlex = False
Else
   chkDoc.Visible = False
End If

If lbPorMoneda Then
    lblTotalC.Width = lblTotalC.Width - (txtHaber.Left - lblTotalC.Left)
    lblTotalC.Left = txtHaber.Left
    lblTotal.Left = lblTotalC.Left + 150
    If gsSimbolo = gcMN Then
       txtDebeD.Visible = False
       txtHaberD.Visible = False
       txtDebe.Left = txtDebeD.Left
       txtHaber.Left = txtHaberD.Left
       fg.Cols = 6
    Else
       txtDebe.Visible = False
       txtHaber.Visible = False
       fg.ColWidth(4) = 0
       fg.ColWidth(5) = 0
    End If
End If

If sMovNroAnt <> "" Then
   cboOperacion.Enabled = False
   txtFecha.Enabled = False
   
   nTipCambioC = gnTipCambio
   GetTipCambio GetFechaMov(sMovNroAnt, True)
   nTipCambio = gnTipCambio
   gnTipCambio = nTipCambioC
   nTipCambioC = gnTipCambioC
   nTipCambioV = gnTipCambioV

   Dim oMov As New DMov
   fgDoc.rsFlex = oMov.CargaMovDocAsiento(nMovNroAnt)
   
   Set rs = oMov.CargaMovOpeAsiento(nMovNroAnt)
   If Not rs.EOF Then
      txtFecha = GetFechaMov(sMovNroAnt, True)
      cboOperacion.AddItem rs!cOpeCod & "  " & rs!cOpeDesc
      cboOperacion.ListIndex = cboOperacion.ListCount - 1
      If Trim(rs!cRuc) = "" Or Trim(rs!cRuc) = "00000000" Then
         txtPersCod.Text = Trim(rs!cPersCod)
      Else
         txtPersCod.Text = Trim(rs!cRuc)
      End If
      txtPersCod.psCodigoPersona = Trim(rs!cPersCod)
      txtPersNombre = PstaNombre(rs!cPersNombre, False)
      txtPersNombre.Tag = rs!cDestino
      txtMovDesc = rs!cMovDesc
   End If
   
   Set rs = oMov.CargaMovCtaAsiento(nMovNroAnt)
   LlenaDatosMovCta
   Set rs = oMov.CargaMovObjAsiento(nMovNroAnt)
   Set oMov = Nothing
   If Not rs Is Nothing Then
      If Not rs.EOF Then
        LlenaDatosMovObj
      End If
   End If
   Sumas
   txtFechaExt = gdFecSis
Else
    txtFecha = gdFecSis
    txtMovDesc = ""
    If lbRegistraPendiente Or lbRegulaPendiente Then
        cboOperacion.AddItem gsOpeCod & "  " & gsOpeDesc
    Else
        If Left(gsOpeCod, 1) = "7" Then
           Set rs = CargaOpeTpo(Mid(gsOpeCod, 1, 4), True, , 0, 2)
        Else
           Set rs = CargaOpeTpo(gsOpeCod, True, , -1)
        End If
        Do While Not rs.EOF
           cboOperacion.AddItem rs!cOpeCod & "  " & rs!cOpeDesc
           rs.MoveNext
        Loop
    End If
   If cboOperacion.ListCount = 1 Then
        cboOperacion.ListIndex = 0
        cboOperacion.Enabled = False
   End If
   fg.TextMatrix(1, 0) = "1"
End If

Set rs = Nothing
If lbExtornar Then
   Me.Caption = "Asientos Contables: Mantenimiento: Extornar"
   cmdNuevo.Visible = False
   cmdEliminar.Visible = False
   lblObj.Visible = False
   fgObj.Visible = False
   txtMovDesc.Enabled = False
   fraMovDesc.Visible = True
   fg.lbEditarFlex = False
   cmdAceptar.Visible = False
   Me.cmdCerrar.Visible = False
   cmdAceptarExt.Visible = True
   cmdCancelarExt.Visible = True
   txtPersCod.Enabled = False
End If

Dim oDoc As New DDocumento
fgDoc.AutoAdd = True
fgDoc.AvanceCeldas = Horizontal
fgDoc.psRaiz = "Documentos"
fgDoc.rsTextBuscar = oDoc.CargaDocumento(, , , 1)
Set oDoc = Nothing
lnNroCta = 0
If lbCargaCuentas Then
   Dim oOpe As New DOperacion
   Dim nPosPendiente As Integer
   Set rs = oOpe.CargaOpeCta(gsOpeCod, , , , , , gOpeCtaCaracObligatorio)
   Set oOpe = Nothing
   'JACA 20110822 se hizo para que tome la cuenta seleccionada
   If gsOpeCod = "741123" Or gsOpeCod = "742123" Then
        If Not (rs.EOF And rs.BOF) Then
            'rs.MoveNext
            fg.AdicionaFila
            n = fg.row
            fg.TextMatrix(n, 0) = n
            fg.TextMatrix(n, 1) = frmAnalisisRegulaPend.sPendiente
            fg.TextMatrix(n, 2) = frmAnalisisRegulaPend.txtCtaPendDes
            nPosPendiente = n

        End If
   Else
     Do While Not rs.EOF
       fg.AdicionaFila
       n = fg.row
       fg.TextMatrix(n, 0) = n
       fg.TextMatrix(n, 1) = rs!cCtaContCod
       fg.TextMatrix(n, 2) = rs!cCtaContDesc
       If lbRegistraPendiente Then
          lnNroCta = lnNroCta + 1
       End If
       If lbRegulaPendiente Then
          If rs!cCtaContCod = frmAnalisisRegulaPend.sPendiente Then
             nPosPendiente = n
          End If
       End If
       rs.MoveNext
    Loop
   End If
   'JACA END
   
   'JACA 20110829 Se cambio de lugar el proceso*****
   If lbRegistraPendiente Or lbRegulaPendiente Then
     CuentaEsPendiente fg.TextMatrix(1, 1), lsClasePendiente
   End If
   'JACA END**************************************
   If lbRegulaPendiente Then
      txtMovDesc = gsGlosa
      txtFecha.Enabled = False
      txtFecha = gdFecha
      If nPosPendiente > 0 Then
        lnSaldo = gnImporte
        If lsClasePendiente = "D" Then
           If Mid(gsOpeCod, 3, 1) = gMonedaExtranjera Then
               fg.TextMatrix(nPosPendiente, 7) = Format(lnSaldo, gsFormatoNumeroView)
           Else
               fg.TextMatrix(nPosPendiente, 5) = Format(lnSaldo, gsFormatoNumeroView)
           End If
        End If
        If lsClasePendiente = "A" Then
           If Mid(gsOpeCod, 3, 1) = gMonedaExtranjera Then
              fg.TextMatrix(nPosPendiente, 6) = Format(lnSaldo, gsFormatoNumeroView)
           Else
              fg.TextMatrix(nPosPendiente, 4) = Format(lnSaldo, gsFormatoNumeroView)
           End If
        End If
      End If
   End If
'Comentado x JACA 20110829 Se cambió de lugar en el proceso***
'   If lbRegistraPendiente Or lbRegulaPendiente Then
'     CuentaEsPendiente fg.TextMatrix(1, 1), lsClasePendiente
'   End If
'JACA END*****************************************************

'   For N = 1 To fg.Rows - 1
'      If fg.TextMatrix(N, 1) <> "" Then
'         Dim oCon As New DConecta
'         oCon.AbreConexion
'         sSql = "SELECT co.cObjetoCod, co.cCtaObjOrden, o.nObjetoNiv, co.nCtaObjNiv, co.cCtaObjFiltro, co.cCtaObjImpre " _
'              & "FROM CtaObj co JOIN Objeto o ON o.cObjetoCod = co.cObjetoCod " _
'              & "WHERE '" & fg.TextMatrix(N, 1) & "' LIKE co.cCtaContCod + '%' "
'         Set rs = oCon.CargaRecordSet(sSql)
'         Do While Not rs.EOF
'            AdicionaObj fg.TextMatrix(N, 0), rs!cCtaObjOrden, "", "", rs!cObjetoCod
'            rs.MoveNext
'         Loop
'         oCon.CierraConexion
'         Set oCon = Nothing
'      End If
'   Next
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not glAceptar Then
   If MsgBox(" ¿ Seguro de Salir sin Grabar ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmación") = vbNo Then
      Cancel = True
   End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub fg_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
Dim sCtaCod  As String
Dim oCtaCont As New DCtaCont
If psDataCod <> "" Then
   fg.TextMatrix(pnRow, 3) = fg.TextMatrix(pnRow, 0)
   Set rs = oCtaCont.CargaCtaObjFiltro(, , "'" & psDataCod & "' Like cCtaContCod+'%' ")
   If Not rs.EOF Then
      sCtaCod = rs!cCtaContCod
   Else
      sCtaCod = psDataCod
   End If
   If Not AsignaCtaObj(sCtaCod, psDataCod) Then
      fg.TextMatrix(pnRow, pnCol) = ""
      fg.TextMatrix(pnRow, pnCol + 1) = ""
   End If
   RefrescaFgObj fg.TextMatrix(fg.row, 0)
End If
Set oCtaCont = Nothing
End Sub
Private Sub RefrescaFgObj(nItem As Integer)
Dim K  As Integer
For K = 1 To fgObj.Rows - 1
    If Len(fgObj.TextMatrix(K, 1)) Then
       If fgObj.TextMatrix(K, 0) = nItem Then
          fgObj.RowHeight(K) = 285
       Else
          fgObj.RowHeight(K) = 0
       End If
    End If
Next
End Sub

Private Sub AdicionaObj(nFila As Integer, psCtaObjOrden As String, psObjetoCod As String, psObjetoDesc As String, psObjCod As String)
Dim nItem As Integer
   fgObj.AdicionaFila
   nItem = fgObj.row
   fgObj.TextMatrix(nItem, 0) = nFila
   fgObj.TextMatrix(nItem, 1) = psCtaObjOrden
   fgObj.TextMatrix(nItem, 2) = psObjetoCod
   fgObj.TextMatrix(nItem, 3) = psObjetoDesc
   fgObj.TextMatrix(nItem, 4) = psObjCod
   fgObj.TextMatrix(nItem, 5) = nFila
   
   
End Sub

Private Sub txtFecha_GotFocus()
txtFecha.SelStart = 0
txtFecha.SelLength = Len(txtFecha.Text)
End Sub

Private Function FechaOk(ByVal sFecha As MaskEdBox, ByVal pbTpoCambio As Boolean) As Boolean
Dim oFun As NContFunciones
FechaOk = True
On Error GoTo FechaOkErr
   If ValidaFecha(sFecha.Text) = "" Then
      If Val(Right(sFecha, 4)) < Year(gdFecSis) - 1 Then
         FechaOk = False
         MsgBox "Año no corresponde a Ejercicio...!", vbInformation, "Error"
         Exit Function
      End If
      Set oFun = New NContFunciones
      If Not PermiteModificarAsiento(oFun.GeneraMovNro(CDate(sFecha), gsCodAge, gsCodUser)) Then
         FechaOk = False
         Set oFun = Nothing
         Exit Function
      End If
      Set oFun = Nothing
      If pbTpoCambio Then
         nTipCambioC = gnTipCambio
         GetTipCambio CDate(sFecha.Text)
         nTipCambio = gnTipCambio
         gnTipCambio = nTipCambioC
         nTipCambioC = gnTipCambioC
         nTipCambioV = gnTipCambioV
      End If
   Else
      MsgBox "Fecha no válida...!", vbInformation, "Error"
      sFecha.SelStart = 0
      sFecha.SelLength = Len(sFecha.Text)
      FechaOk = False
   End If
Exit Function
FechaOkErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Function
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If FechaOk(txtFecha, True) Then
      If txtPersCod.Enabled Then
         txtPersCod.SetFocus
      ElseIf txtMovDesc.Enabled Then
         txtMovDesc.SetFocus
      Else
         fg.SetFocus
      End If
   End If
End If
End Sub

Private Sub txtFecha_Validate(Cancel As Boolean)
If Not FechaOk(txtFecha, True) Then
   Cancel = True
End If
End Sub

Private Sub txtFechaExt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If FechaOk(txtFechaExt, False) Then
      txtMovDescExt.SetFocus
   End If
End If
End Sub

Private Sub txtFechaExt_Validate(Cancel As Boolean)
If Not FechaOk(txtFechaExt, False) Then
   Cancel = True
End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   fg.SetFocus
End If
End Sub

Private Function AsignaCtaObj(ByVal psCtaContCod As String, ByVal psCtaDato As String) As Boolean
Dim lsFiltro As String
Dim lsRaiz   As String
Dim lsSubCta As String
Dim lbLeeDatos As Boolean

Dim rs       As ADODB.Recordset
Dim rsCtaObj As ADODB.Recordset
Dim rsFiltro As ADODB.Recordset

Dim oDescObj As ClassDescObjeto
Dim UP       As UPersona
Dim oRHAreas As DActualizaDatosArea
Dim oIF      As NCajaCtaIF
Dim oEfectiv As Defectivo

Dim nPosSubCta As Integer
Set oDescObj = New ClassDescObjeto
Set UP = New UPersona
Set oRHAreas = New DActualizaDatosArea
Set oContFunct = New NContFunciones
Set oIF = New NCajaCtaIF
Set oEfectiv = New Defectivo

Set rs = New ADODB.Recordset
Set rsCtaObj = New ADODB.Recordset
AsignaCtaObj = True
oDescObj.lbUltNivel = True

EliminaFgObj Val(fg.TextMatrix(fg.row, 0))
nPosSubCta = 1
Set rsCtaObj = oCta.CargaCtaObj(psCtaContCod)
If Not rsCtaObj.EOF And Not rsCtaObj.BOF Then
   Do While Not rsCtaObj.EOF
      lbLeeDatos = False
      lsSubCta = Mid(psCtaDato, nPosSubCta + Len(psCtaContCod), Len(psCtaDato) - Len(psCtaContCod))
      Set rsFiltro = oContFunct.GetSubCtaContFiltro(psCtaContCod, rsCtaObj!cObjetoCod, lsSubCta, rsCtaObj!cCtaObjFiltro)
      If Not rsFiltro Is Nothing Then
         If Not rsFiltro.EOF Then
            If rsFiltro.RecordCount = 1 Then
               nPosSubCta = nPosSubCta + Len(rsFiltro!cSubCtaCod)
               AdicionaObj fg.TextMatrix(fg.row, 0), rsCtaObj!nCtaObjOrden, rsFiltro!cObjetoCod, _
                           rsFiltro!cObjetoDesc, rsCtaObj!cObjetoCod
            Else
               nPosSubCta = nPosSubCta + Len(rsFiltro!cSubCtaCod)
               oDescObj.Show rsFiltro, "", "Objetos"
               If oDescObj.lbOk Then
                  AdicionaObj fg.TextMatrix(fg.row, 0), rsCtaObj!nCtaObjOrden, oDescObj.gsSelecCod, _
                               oDescObj.gsSelecDesc, rsCtaObj!cObjetoCod
               Else
                  EliminaFgObj Val(fg.TextMatrix(fg.row, 0))
                  AsignaCtaObj = False
                  GoTo AsignaCtaObjSalir
               End If
            End If
         Else
            lbLeeDatos = True
         End If
      Else
         lbLeeDatos = True
      End If
      If lbLeeDatos Then
         lsRaiz = ""
         lsFiltro = ""
         Select Case Val(rsCtaObj!cObjetoCod)
             Case ObjCMACAgencias
                 Set rs = oRHAreas.GetAgencias
             Case ObjCMACAgenciaArea
                 lsRaiz = "Unidades Organizacionales"
                 Set rs = oRHAreas.GetAgenciasAreas
             Case ObjCMACArea
                 Set rs = oRHAreas.GetAreas
             Case ObjEntidadesFinancieras
                 Set rs = oIF.GetCtasInstFinancieras(rsCtaObj!cCtaObjFiltro)
             Case ObjDescomEfectivo
                 lsRaiz = "COMPOSICION DE EFECTIVO"
                 Set rs = oEfectiv.GetBilletajes(rsCtaObj!cCtaObjFiltro)
             Case ObjPersona
                 Set rs = Nothing
             Case Else
                 Set rs = GetObjetos(Val(rsCtaObj!cObjetoCod))
         End Select
         If Not rs Is Nothing Then
            If Not rs.EOF And Not rs.BOF Then
               If rs.RecordCount > 1 Then
                  oDescObj.Show rs, "", lsRaiz
                  If oDescObj.lbOk Then
                     lsFiltro = oContFunct.GetFiltroObjetos(Val(rsCtaObj!cObjetoCod), psCtaContCod, oDescObj.gsSelecCod, False)
                     nPosSubCta = nPosSubCta + Len(lsFiltro)
                     AdicionaObj fg.TextMatrix(fg.row, 0), rsCtaObj!nCtaObjOrden, oDescObj.gsSelecCod, _
                                 oDescObj.gsSelecDesc, rsCtaObj!cObjetoCod
                  Else
                     AsignaCtaObj = False
                     GoTo AsignaCtaObjSalir
                  End If
               Else
                  nPosSubCta = nPosSubCta + Len(rs!cSubCtaCod)
                  AdicionaObj fg.TextMatrix(fg.row, 0), rsCtaObj!nCtaObjOrden, rsCtaObj!cObjetoCod, _
                              rsCtaObj!cObjetoDesc, rsCtaObj!cObjetoCod
               End If
            End If
         Else
            If Val(rsCtaObj!cObjetoCod) = ObjPersona Then
               If sPersona = "" Then
                  Set UP = frmBuscaPersona.Inicio
                  If Not UP Is Nothing Then
                     AdicionaObj fg.TextMatrix(fg.row, 0), rsCtaObj!nCtaObjOrden, _
                                       UP.sPersCod, UP.sPersNombre, _
                                       rsCtaObj!cObjetoCod
                  Else
                     AsignaCtaObj = False
                     GoTo AsignaCtaObjSalir
                  End If
               Else
                  AdicionaObj fg.TextMatrix(fg.row, 0), rsCtaObj!cCtaObjOrde, sPersona, sNomPers, rsCtaObj!cObjetoCod
               End If
            End If
         End If
      End If
      rsCtaObj.MoveNext
   Loop
End If
AsignaCtaObjSalir:
RSClose rsCtaObj
Set oDescObj = Nothing
Set UP = Nothing
Set oRHAreas = Nothing
Set oContFunct = Nothing
End Function

Private Sub EliminaFgObj(nItem As Integer)
Dim K  As Integer
K = 1
Do While K < fgObj.Rows
   If Len(fgObj.TextMatrix(K, 1)) > 0 Then
      If Val(fgObj.TextMatrix(K, 0)) = nItem Then
         If fgObj.TextMatrix(K, 2) = Format(ObjPersona, "##") Then
            sPersona = fgObj.TextMatrix(K, 2)
            sNomPers = fgObj.TextMatrix(K, 3)
         End If
         fgObj.EliminaFila K, False
      Else
         K = K + 1
      End If
   Else
      K = K + 1
   End If
Loop
End Sub

Private Sub LlenaDatosMovCta()
Dim nRow As Double
Do While Not rs.EOF
   With fg
      .AdicionaFila
      nRow = .row
      .TextMatrix(nRow, 0) = rs!nMovItem
      .TextMatrix(nRow, 1) = rs!cCtaContCod
      .TextMatrix(nRow, 2) = rs!cCtaContDesc
      .TextMatrix(nRow, 3) = rs!nMovItem
      .TextMatrix(nRow, 4) = rs!nDebe
      .TextMatrix(nRow, 5) = rs!nHaber
      .TextMatrix(nRow, 6) = rs!nDebeME
      .TextMatrix(nRow, 7) = rs!nHaberME
   End With
   rs.MoveNext
Loop
End Sub

Private Sub LlenaDatosMovObj()
Dim nRow As Double
Dim prsObj As ADODB.Recordset
rs.MoveFirst
Do While Not rs.EOF
   With fgObj
      .AdicionaFila
      nRow = .row
      .TextMatrix(nRow, 0) = Val(rs!nMovItem)
      .TextMatrix(nRow, 1) = rs!nMovObjOrden
      .TextMatrix(nRow, 2) = rs!cObjetoCod
      .TextMatrix(nRow, 3) = rs!cObjetoDesc
      .TextMatrix(nRow, 4) = rs!ObjPadre
      .TextMatrix(nRow, 5) = rs!nMovItem
   End With
   rs.MoveNext
Loop
RSClose prsObj
End Sub

Private Sub txtMovDescExt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   KeyAscii = 0
   cmdAceptarExt.SetFocus
End If
End Sub

Private Sub txtPersCod_EmiteDatos()
txtPersNombre = txtPersCod.psDescripcion
If txtPersNombre <> "" Then
   txtMovDesc.SetFocus
End If
End Sub

Public Property Get lOk() As Boolean
lOk = OK
End Property

Public Property Let lOk(ByVal vNewValue As Boolean)
OK = vNewValue
End Property

