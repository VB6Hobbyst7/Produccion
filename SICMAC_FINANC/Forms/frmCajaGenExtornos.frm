VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenExtornos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5820
   ClientLeft      =   1470
   ClientTop       =   2265
   ClientWidth     =   10395
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajaGenExtornos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkConGenAsiento 
      Caption         =   "Extorno con Generacion de Asiento"
      Height          =   210
      Left            =   6705
      TabIndex        =   28
      Top             =   5580
      Width           =   3540
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Nueva Fecha"
      ForeColor       =   &H8000000D&
      Height          =   675
      Left            =   60
      TabIndex        =   24
      Top             =   4890
      Visible         =   0   'False
      Width           =   3555
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         Height          =   315
         Left            =   1290
         TabIndex        =   26
         ToolTipText     =   "Aplicar cambio de Fecha de Asiento"
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2370
         TabIndex        =   25
         ToolTipText     =   "Cancelar Cambio de Fecha de Asiento"
         Top             =   240
         Width           =   1035
      End
      Begin MSMask.MaskEdBox txtMovFecha 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
   End
   Begin VB.CommandButton cmdCambiarF 
      Caption         =   "&Cambiar Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   45
      TabIndex        =   9
      Top             =   30
      Width           =   10290
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
         Left            =   8805
         TabIndex        =   2
         Top             =   203
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtDesde 
         Height          =   345
         Left            =   2895
         TabIndex        =   0
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
         Left            =   4605
         TabIndex        =   1
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
      Begin VB.Frame FraFechaMov 
         Height          =   495
         Left            =   6540
         TabIndex        =   17
         Top             =   135
         Visible         =   0   'False
         Width           =   2040
         Begin MSMask.MaskEdBox txtFechaMov 
            Height          =   300
            Left            =   870
            TabIndex        =   20
            Top             =   150
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            ForeColor       =   -2147483635
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
            AutoSize        =   -1  'True
            Caption         =   "Mov. Al..."
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
            Left            =   90
            TabIndex        =   19
            Top             =   180
            Width           =   720
         End
      End
      Begin VB.Label lbltitulo 
         AutoSize        =   -1  'True
         Caption         =   "EXTORNOS"
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
         Left            =   210
         TabIndex        =   12
         Top             =   255
         Width           =   1290
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
         Left            =   4005
         TabIndex        =   11
         Top             =   300
         Width           =   540
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
         Left            =   2235
         TabIndex        =   10
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.Frame FraLista 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Left            =   45
      TabIndex        =   8
      Top             =   720
      Width           =   10290
      Begin Sicmact.TxtBuscar txtCuenta 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   4305
         Visible         =   0   'False
         Width           =   1965
         _ExtentX        =   3466
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
      Begin VB.TextBox txtSubCuenta 
         Height          =   345
         Left            =   5310
         MaxLength       =   10
         TabIndex        =   7
         Top             =   4320
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   18
         Top             =   4290
         Width           =   1470
      End
      Begin VB.TextBox txtMovDesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1215
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3750
         Width           =   8940
      End
      Begin VB.TextBox txtConcepto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   165
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2925
         Width           =   9990
      End
      Begin Sicmact.FlexEdit fgListaCG 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4683
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Fecha-Institución Origen-Institución Fuente-Importe-cMovNro"
         EncabezadosAnchos=   "350-1000-3500-3500-1200-2500"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-R-L"
         FormatosEdit    =   "0-0-0-0-2-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "&Extornar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1100
         TabIndex        =   15
         Top             =   4320
         Width           =   1470
      End
      Begin VB.CommandButton cmdConfRetiro 
         Caption         =   "&Confirmar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7120
         TabIndex        =   13
         Top             =   4290
         Width           =   1470
      End
      Begin VB.CommandButton cmdConfAper 
         Caption         =   "&Confirmar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7140
         TabIndex        =   16
         Top             =   4290
         Width           =   1470
      End
      Begin VB.Label lblCuenta 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable"
         Height          =   210
         Left            =   210
         TabIndex        =   23
         Top             =   4410
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblSubCuenta 
         AutoSize        =   -1  'True
         Caption         =   "Sub Cuenta"
         Height          =   210
         Left            =   4260
         TabIndex        =   22
         Top             =   4410
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción"
         Height          =   315
         Left            =   210
         TabIndex        =   21
         Top             =   3810
         Width           =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   10260
         X2              =   -15
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   10260
         X2              =   -15
         Y1              =   3570
         Y2              =   3570
      End
   End
End
Attribute VB_Name = "frmCajaGenExtornos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCaja As nCajaGeneral
Dim objPista As COMManejador.Pista
Dim oAdeud As NCajaAdeudados

'***Modificado por ELRO el 20120103, según Acta Acta 311-2011/TI-D
Dim ldFecCie As Date
'***Fin Modificado por ELRO**********************************

Private Sub cmdAplicar_Click()
Dim nPos       As Variant
Dim sMovNro    As String
Dim sMovCambio As String
Dim sFechaNew  As String
Dim oFun As New NContFunciones
   If MsgBox(" ¿ Seguro de Modificar Fecha de Movimiento ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
      If Month(txtMovFecha) <> nVal(Mid(fgListaCG.TextMatrix(fgListaCG.row, 12), 5, 2)) And Mid(gsOpeCod, 3, 1) = 2 Then
         MsgBox "No se puede cambiar Asientos de Moneda Extranjera a un mes diferente", vbInformation, "¡Aviso!"
         Exit Sub
      End If
      sMovNro = oFun.GeneraMovNro(CDate(txtMovFecha), , , Format(txtMovFecha.Text, gsFormatoMovFecha) & Mid(fgListaCG.TextMatrix(fgListaCG.row, 12), 9, 25))
      sMovCambio = oFun.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
      sFechaNew = Format(txtMovFecha, gsFormatoFecha)
      If Not oFun.PermiteModificarAsiento(sMovNro, False) Then
         MsgBox "No de puede cambiar fecha de Asiento a un mes ya Cerrado", vbInformation, "Aviso"
         Exit Sub
      End If
      
      Dim oMov As New DMov
      oMov.BeginTrans
      oMov.ActualizaMovimiento sMovNro, fgListaCG.TextMatrix(fgListaCG.row, 12)
      oMov.InsertaMovModifica sMovCambio, fgListaCG.TextMatrix(fgListaCG.row, 12), sMovNro
      oMov.ActualizaCtaIF fgListaCG.TextMatrix(fgListaCG.row, 12), fgListaCG.TextMatrix(fgListaCG.row, 12), fgListaCG.TextMatrix(fgListaCG.row, 12), , sFechaNew, sFechaNew, sFechaNew, , sFechaNew, , , Format(Me.txtMovFecha, gsFormatoMovFecha) & Mid(fgListaCG.TextMatrix(fgListaCG.row, 12), 9, 24)
      oMov.ActualizaCtaIFAdeudado fgListaCG.TextMatrix(fgListaCG.row, 12), fgListaCG.TextMatrix(fgListaCG.row, 12), fgListaCG.TextMatrix(fgListaCG.row, 12), , , , , , , , , , , , sFechaNew
      oMov.CommitTrans
      Set oMov = Nothing
   End If
   Set oFun = Nothing
   OpcionesCambiar False
End Sub

Private Sub OpcionesCambiar(lOp As Boolean)
fraFecha.Visible = lOp
cmdCambiarF.Visible = Not lOp
cmdExtornar.Visible = Not lOp
txtMovFecha.Visible = lOp
End Sub

Private Sub cmdCambiarF_Click()
Dim oCont As New NContFunciones
On Error GoTo CambiarErr
If Me.fgListaCG.TextMatrix(1, 1) = "" Then
    Exit Sub
End If
If Not oCont.PermiteModificarAsiento(Me.fgListaCG.TextMatrix(fgListaCG.row, 12)) Then
   Exit Sub
End If
OpcionesCambiar True
txtMovFecha.SetFocus
Exit Sub
CambiarErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdCancelar_Click()
   OpcionesCambiar False
End Sub

Private Sub cmdConfAper_Click()
Dim lnImporte   As Currency
Dim lsPersCod   As String
Dim lsTipoIF    As String
Dim lsCtaIFCod  As String
Dim lnMovRef    As Double

If Validar() = False Then Exit Sub

'cmdConfAper.Enabled = False
lnImporte = fgListaCG.TextMatrix(fgListaCG.row, 5)
lsPersCod = fgListaCG.TextMatrix(fgListaCG.row, 11)
lsTipoIF = fgListaCG.TextMatrix(fgListaCG.row, 10)
lsCtaIFCod = fgListaCG.TextMatrix(fgListaCG.row, 8)
lnMovRef = fgListaCG.TextMatrix(fgListaCG.row, 6)

If gsOpeCod = gOpeCGAdeudaRegPagareConfMN Or gsOpeCod = gOpeCGAdeudaRegPagareConfMe Then
    ConfirmaAdeudados lnImporte, lsPersCod, lsTipoIF, lsCtaIFCod, lnMovRef
Else
    ConfirmaOtros lnImporte, lsPersCod, lsTipoIF, lsCtaIFCod, lnMovRef
End If
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
        Set objPista = Nothing
        '****
Exit Sub
ConfirmaAperErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub cmdConfRetiro_Click()
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim lnMovRef As Long
Dim lnMonto As Currency
Dim ldFechaMov As Date
Dim rsBill As ADODB.Recordset
Dim rsMon As ADODB.Recordset

Dim oCont As NContFunciones
Dim oOpe As DOperacion
Dim lsPersCod  As String
Dim lsIFTpo    As String
Dim lsCtaIFCod As String

On Error GoTo ConfRetitoErr

Set oCont = New NContFunciones
Set oOpe = New DOperacion

If fgListaCG.TextMatrix(1, 0) = "" Then Exit Sub

Set rsBill = New ADODB.Recordset
Set rsMon = New ADODB.Recordset
lnMonto = fgListaCG.TextMatrix(fgListaCG.row, 6)
If lnMonto = 0 Then
    MsgBox "Monto no válido para Operación", vbInformation, "Aviso"
    Exit Sub
End If
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción de Operación no ingresada", vbInformation, "aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If

'***Modificado por ELRO el 20110930, según Acta 269-2011/TI-D y Acta 311-2011/TI-D
If gdFecSis <= ldFecCie Then
    MsgBox "Mes ya Cerrado. Imposible realizar la operación...!", vbInformation, "Aviso"
    Exit Sub
End If
'***Fin Modificado por ELRO*******************************************************

frmCajaGenEfectivo.Inicio gsOpeCod, gsOpeDesc, lnMonto, Mid(gsOpeCod, 3, 1), True
If frmCajaGenEfectivo.lbOk Then
    Set rsBill = frmCajaGenEfectivo.rsBilletes
    Set rsMon = frmCajaGenEfectivo.rsMonedas
    ldFechaMov = frmCajaGenEfectivo.FechaMov
Else
    Set rsBill = Nothing
    Set rsMon = Nothing
    Exit Sub
End If
If MsgBox("Desea Confirmar el retiro seleccionado", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lnMovRef = fgListaCG.TextMatrix(fgListaCG.row, 7)
    lsPersCod = fgListaCG.TextMatrix(fgListaCG.row, 9)
    lsIFTpo = fgListaCG.TextMatrix(fgListaCG.row, 10)
    lsCtaIFCod = fgListaCG.TextMatrix(fgListaCG.row, 11)
    
    lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D")
    lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
    
    gsMovNro = oCont.GeneraMovNro(ldFechaMov, gsCodAge, gsCodUser)
    oCaja.GrabaMovEfectivo gsMovNro, gsOpeCod, txtMovDesc, _
            rsBill, rsMon, lsCtaHaber, lsCtaDebe, lnMonto, ObjEntidadesFinancieras, lsIFTpo & "." & lsPersCod & "." & lsCtaIFCod, -1, "", gdFecSis, True, lnMovRef
    
    ImprimeAsientoContable gsMovNro, , , , True, True, txtMovDesc
    
    Unload frmCajaGenEfectivo
    Set frmCajaGenEfectivo = Nothing
    objPista.InsertarPista gsOpeCod, gsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Confirmacion de Retiro"
    If MsgBox("Desea realizar otro Movimiento de Confirmación??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        fgListaCG.EliminaFila fgListaCG.row
        fgListaCG.SetFocus
        txtMovDesc = ""
        txtConcepto = ""
    Else
        Unload Me
    End If
End If
Exit Sub
ConfRetitoErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdExtornar_Click()
Dim oCon As NContFunciones
Dim oOpe As DOperacion
Dim lnMovNro As String
Dim lsMovNroExt As String
Dim lnNumTran As Long
Dim lnImporte As Currency
Dim ldFechaMov As Date
Dim lbEliminaMov As Boolean
Dim lsMovNro As String
Dim lsAgeCodRef As String
Dim lsDocNro As String
Dim lsObjetoCod As String
Dim lsAreaCod As String
Dim lsAgeCod  As String
Dim ldFecReg  As Date
Dim oCaja As New nCajaGeneral
Dim oCajaIF As New DCajaCtasIF
Dim lsPersCod As String
Dim lsIFTpo As String
Dim lsCtaIFCod As String
Dim lnNroCuota As Integer
On Error GoTo ExtornarErr

Set oCon = New NContFunciones
Set oOpe = New DOperacion

'***Agregado por ELRO el 20120109, según Acta N° 003-2012/TI-D
Dim dFechaCierreMensualContabilidad, dFechaHabil As Date
Dim i, nDias As Integer
Dim oNConstSistemas As NConstSistemas
Set oNConstSistemas = New NConstSistemas
Dim oNContFunciones As New NContFunciones
Set oNContFunciones = New NContFunciones
  
dFechaCierreMensualContabilidad = CDate(oNConstSistemas.LeeConstSistema(gConstSistCierreMensualCont))
'***Fin Agregado por ELRO*************************************

If fgListaCG.TextMatrix(1, 1) = "" Then
    MsgBox "No existen Movimientos para Extornar", vbInformation, "¡Aviso!"
    Exit Sub
End If

If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Falta indicar motivo de Extorno", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If

lnMovNro = 0
lsDocNro = ""
lsObjetoCod = ""
Select Case gsOpeCod
    Case gOpeMEExtCompraAInst, gOpeMEExtVentaAInst
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 9)
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 7)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 6))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
    Case gOpeMEExtCompraEfect, gOpeMEExtVentaEfec
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 6)
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 7)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 3))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
    '*******************************
    '** Modify GITU 29-09-2008
    Case gOpeCGTransfExtBancosMN, gOpeCGTransfExtCMACSBancosMN, gOpeCGTransfExtBancosCMACSMN, gOpeCGTransfExtMismoBancoMN, "401435", _
        gOpeCGTransfExtBancosME, gOpeCGTransfExtCMACSBancosME, gOpeCGTransfExtBancosCMACSME, gOpeCGTransfExtMismoBancoME, "402435"
    '*******************************
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 9)
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 7)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 6))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
    Case gOpeCGRVentanaIngresoMNExt, gOpeCGRVentanaIngresoMEExt
        If nVal(fgListaCG.TextMatrix(fgListaCG.row, 3)) = 0 Or Trim(fgListaCG.TextMatrix(fgListaCG.row, 9)) = "" Then
            MsgBox "Operación no se puede extornar. Imposible determinar Agencia donde se realizo se ingreso", vbInformation, "¡Aviso!"
            Exit Sub
        End If
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 8)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 7)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 4))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
        lsAgeCodRef = Trim(fgListaCG.TextMatrix(fgListaCG.row, 9))
        lnNumTran = fgListaCG.TextMatrix(fgListaCG.row, 3)

    Case gOpeCGExtBcoDepEfectivo, gOpeCGExtBcoDepEfectivoME, _
         gOpeCGExtBcoRetEfectivo, gOpeCGExtBcoRetEfectivoME, _
         gOpeCGExtBcoConfRetEfectivo, gOpeCGExtBcoConfRetEfectivoME, _
         gOpeCGExtBcoGastComision, gOpeCGExtBcoGastComisionME, _
         gOpeCGExtBcoIntCtasAho, gOpeCGExtBcoIntCtasAhoME, _
         gOpeCGOpeCMACDepDivMNExt, gOpeCGOpeCMACDepDivMEExt, _
         gOpeCGOpeCMACRetDivMNExt, gOpeCGOpeCMACRetDivMEExt, _
         gOpeCGOpeCMACRegularizMNExt, gOpeCGOpeCMACRegularizMEExt, _
         gOpeCGOpeCMACAperCtasMNExt, gOpeCGOpeCMACAperCtasMEExt, _
         gOpeCGOpeCMACExtDepEfeMN, gOpeCGOpeCMACExtDepEfeME, _
         gOpeCGOpeCMACExtRetEfeMN, gOpeCGOpeCMACExtRetEfeME, _
         gOpeCGOpeCMACExtConRetEfeMN, gOpeCGOpeCMACExtConRetEfeME
         
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 12)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 11)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 6))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
        
    Case gOpeCGExtBcoRetDiv, gOpeCGExtBcoRetDivME, _
         gOpeCGExtBcoDepDiv, gOpeCGExtBcoDepDivME, _
         gOpeCGOpeBancosOtrosDepositosMNExt, gOpeCGOpeBancosOtrosRetirosMNExt, _
         gOpeCGOpeBancosOtrosDepositosMEExt, gOpeCGOpeBancosOtrosRetirosMEExt
         
         
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 12)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 11)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 6))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
        lsAgeCodRef = Trim(fgListaCG.TextMatrix(fgListaCG.row, 15))
        lnNumTran = fgListaCG.TextMatrix(fgListaCG.row, 14)
        
    Case gOpeCGExtBcoRegCheques, gOpeCGExtBcoRegChequesME, _
         gOpeCGExtBcoDepCheques, gOpeCGExtBcoDepChequesME, _
         gOpeCGExtBcoRecepChqRegAgencias, gOpeCGExtBcoRecepChqRegAgenciasME
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 12)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 11)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 6))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
        lsDocNro = fgListaCG.TextMatrix(fgListaCG.row, 5)
        lsObjetoCod = fgListaCG.TextMatrix(fgListaCG.row, 9) + "." + fgListaCG.TextMatrix(fgListaCG.row, 8) + "." + fgListaCG.TextMatrix(fgListaCG.row, 3)
        lsAgeCodRef = fgListaCG.TextMatrix(fgListaCG.row, 15)
        If fgListaCG.Cols > 16 Then
            ldFecReg = fgListaCG.TextMatrix(fgListaCG.row, 16)
        End If
    Case gOpeCGExtBcoApertCta, gOpeCGExtBcoApertCtaME, _
         gOpeCGExtBcoConfApert, gOpeCGExtBcoConfApertME, _
         gOpeCGExtBcoIntDevengPF, gOpeCGExtBcoIntDevengPFME, _
         gOpeCGExtBcoCapitalizaIntDPF, gOpeCGExtBcoCapitalizaIntDPFME, _
         gOpeCGExtBcoCancelaCtas, gOpeCGExtBcoCancelaCtasME, _
         gOpeCGOpeCMACConfAperMNExt, gOpeCGOpeCMACConfAperMEExt, _
         gOpeCGOpeCMACIntDevPFMNExt, gOpeCGOpeCMACIntDevPFMEExt, _
         gOpeCGOpeCMACGastosComMNExt, gOpeCGOpeCMACGastosComMEExt, _
         gOpeCGOpeCMACCapIntDevPFMNExt, gOpeCGOpeCMACCapIntDevPFMEExt, _
         gOpeCGOpeCMACCancelaMNExt, gOpeCGOpeCMACCancelaMEExt, _
         gOpeCGOpeCMACInteresAhoMNExt, gOpeCGOpeCMACInteresAhoMEExt
        
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 12)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 11)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 6))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
        lsObjetoCod = fgListaCG.TextMatrix(fgListaCG.row, 9) + "." + fgListaCG.TextMatrix(fgListaCG.row, 8) + "." + fgListaCG.TextMatrix(fgListaCG.row, 10)
         
    Case gOpeCGAdeudaExtRegistroMN, gOpeCGAdeudaExtRegistroME, _
         gOpeCGAdeudaExtConfRegiMN, gOpeCGAdeudaExtConfRegiME
        
        lnImporte = fgListaCG.TextMatrix(fgListaCG.row, 5)
        lsObjetoCod = fgListaCG.TextMatrix(fgListaCG.row, 10) & "." & fgListaCG.TextMatrix(fgListaCG.row, 11) & "." & fgListaCG.TextMatrix(fgListaCG.row, 8)
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 6)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 12)
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
        
    Case gOpeCGAdeudaExtProvisiónMN, gOpeCGAdeudaExtProvisiónME, _
         gOpeCGAdeudaExtPagoCuotaMN, gOpeCGAdeudaExtPagoCuotaME, _
         gOpeCGAdeudaExtReprogramaMN, gOpeCGAdeudaExtReprogramaME
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 10)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 9)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 4))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
        
        If gsOpeCod = gOpeCGAdeudaExtProvisiónMN Or gsOpeCod = gOpeCGAdeudaExtProvisiónME Then
            lnNroCuota = fgListaCG.TextMatrix(fgListaCG.row, 12)
            lsPersCod = fgListaCG.TextMatrix(fgListaCG.row, 13)
            lsIFTpo = fgListaCG.TextMatrix(fgListaCG.row, 14)
            lsCtaIFCod = fgListaCG.TextMatrix(fgListaCG.row, 15)
        ElseIf gsOpeCod = gOpeCGAdeudaExtPagoCuotaMN Or gsOpeCod = gOpeCGAdeudaExtPagoCuotaME Then
            lnNroCuota = fgListaCG.TextMatrix(fgListaCG.row, 12)
            lsPersCod = fgListaCG.TextMatrix(fgListaCG.row, 13)
            lsIFTpo = fgListaCG.TextMatrix(fgListaCG.row, 14)
            lsCtaIFCod = fgListaCG.TextMatrix(fgListaCG.row, 15)
        End If
    Case OpeCGCartaFianzaIngExt, _
         OpeCGCartaFianzaIngMEExt
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 11)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 10)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 5))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 9))

    Case OpeCGOtrosOpeEfecIngrExt, OpeCGOtrosOpeEfecEgreExt, _
         OpeCGOtrosOpeEfecCambExt, OpeCGOtrosOpeEfecOtroExt, _
         OpeCGOtrosOpeEfecIngrMEExt, OpeCGOtrosOpeEfecEgreMEExt, _
         OpeCGOtrosOpeEfecCambMEExt, OpeCGOtrosOpeEfecOtroMEExt
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 9)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 8)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 4))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
        
    Case gOpeCGExtProvPagoEfectivo, gOpeCGExtProvPagoTransfer, _
         gOpeCGExtProvPagoOPago, gOpeCGExtProvPagoCheque, _
         gOpeCGExtProvPagoAbono, gOpeCGExtProvPagoRechazo, _
         gOpeCGExtProvEntrOPago, gOpeCGExtProvEntrCheques, _
         gOpeCGExtProvPagoEfectivoME, gOpeCGExtProvPagoTransferME, _
         gOpeCGExtProvPagoOPagoME, gOpeCGExtProvPagoChequeME, _
         gOpeCGExtProvPagoAbonoME, gOpeCGExtProvPagoRechazoME, _
         gOpeCGExtProvEntrOPagoME, gOpeCGExtProvEntrChequesME
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 9)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 8)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 4))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
        
        If gsOpeCod = gOpeCGExtProvPagoAbono Or gsOpeCod = gOpeCGExtProvPagoAbonoME Then 'PASIERS1242014
             If Not (ldFechaMov = gdFecSis) Then
                MsgBox "El extorno de la operación sólo se puede realizar en el mismo dia que se realiza la operación de depósito.", vbInformation, "¡Aviso!"
                Exit Sub
            End If
        End If 'END PASI
    Case gOpeCGExtProvDevSUNAT 'PASIERS1242014
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 9)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 8)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 4))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
        If Not (ldFechaMov = gdFecSis) Then
                MsgBox "El extorno de la operación sólo se puede realizar en el mismo dia que se realiza la operación de depósito.", vbInformation, "¡Aviso!"
                Exit Sub
        End If
    Case OpeCGOtrosOpeRetPagSeguroDesgravamenMNExt, OpeCGOtrosOpeRetPagSeguroIncendioMNExt, _
            OpeCGOtrosOpeRetPagSeguroDesgravamenMEExt, OpeCGOtrosOpeRetPagSeguroIncendioMEExt 'PASIERS1362014
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 9)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 8)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 6))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
    'END PASI
    
    Case OpeCGCartaFianzaSalExt, OpeCGCartaFianzaSalMEExt
        Dim rs As New ADODB.Recordset
        
        Set rs = oCaja.GetMovRefCartaFianza(fgListaCG.TextMatrix(fgListaCG.row, 11))
        lnMovNro = rs!nMovNro
        lsMovNro = rs!cMovNro 'fgListaCG.TextMatrix(fgListaCG.Row, 10)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 5))
        ldFechaMov = Right(Left(rs!cMovNro, 8), 2) & "/" & Mid(Left(rs!cMovNro, 8), 5, 2) & "/" & Left(rs!cMovNro, 4)
        
    Case Else
        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 9)
        lsMovNro = fgListaCG.TextMatrix(fgListaCG.row, 8)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 4))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
End Select

If lnMovNro = 0 Or lsMovNro = "" Then
    MsgBox "Aún no esta implementado Extorno de esta Operación. Consultar con Sistemas", vbInformation, "¡Aviso!"
    Exit Sub
End If

If MsgBox("Desea Realizar el Extorno respectivo??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
   Dim oFun As New NContFunciones
   lbEliminaMov = oFun.PermiteModificarAsiento(lsMovNro, False)
      
   '***Modificado por ELRO el 20111114 según Acta 311-2011/TI-D
   'If Not lbEliminaMov Then
   'If MsgBox("Fecha de Extorno corresponde a un mes ya Cerrado, ¿ Desea Extornar Operación ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then Exit Sub
   'End If
   If gsOpeCod = gOpeCGOpeCMACExtDepEfeMN Or _
        gsOpeCod = gOpeCGOpeCMACExtDepEfeME Or _
        gsOpeCod = gOpeCGOpeCMACExtRetEfeMN Or _
        gsOpeCod = gOpeCGOpeCMACExtRetEfeME Or _
        gsOpeCod = gOpeCGOpeCMACExtConRetEfeMN Or _
        gsOpeCod = gOpeCGOpeCMACExtConRetEfeME Then
        If Not lbEliminaMov Then
            MsgBox "Mes ya Cerrado. Imposible realizar la operación...!", vbInformation, "Aviso"
            Exit Sub
        Else
            If CDate(txtFechaMov) < gdFecSis Then
                MsgBox "Día ya Cerrado. Imposible realizar la operación...!", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    Else

         If CInt(Right(Left(lsMovNro, 6), 2)) = CInt(Month(dFechaCierreMensualContabilidad)) And _
             CInt(Left(lsMovNro, 4)) = CInt(Year(dFechaCierreMensualContabilidad)) Then
             
             If MsgBox("¿Desea realizar la operación en una fecha que pertenece a un Mes Cerrado?", vbYesNo, "Confirmar") = vbYes Then
             nDias = DateDiff("D", dFechaCierreMensualContabilidad, gdFecSis)
             
                 For i = 1 To nDias
                 
                    If Not oNContFunciones.EsFeriado(DateAdd("D", i, dFechaCierreMensualContabilidad)) Then
                        dFechaHabil = DateAdd("D", i, dFechaCierreMensualContabilidad)
                        If DateDiff("D", dFechaHabil, gdFecSis) > 0 Then
                            MsgBox "Solo se puede realizar la operación en un Mes Cerrado hasta " & dFechaHabil, vbInformation, "aviso"
                            Exit Sub
                            
                        End If
                    
                    End If
                  
                 Next i
            Else
                Exit Sub
            End If
         Else
             If Not lbEliminaMov Then
                 '***Modificado por ELRO el 20120110, según Acta N° 003-2012/TI-D
                 'If MsgBox("Fecha de Extorno corresponde a un mes ya Cerrado, ¿ Desea Extornar Operación ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then Exit Sub
                 MsgBox "Fecha de Extorno corresponde a un Mes Cerrado"
                 cmdExtornar.SetFocus
                 Exit Sub
                 '***Fin Modificado por ELRO*************************************
             End If
         End If
    End If
   '***Fin Modificado por ELRO*********************************
   
   Set oFun = Nothing
   cmdExtornar.Enabled = False

   If lbEliminaMov Then
        If Me.chkConGenAsiento.value = 1 Then
            lbEliminaMov = False
        End If
   End If

    lsMovNroExt = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    'ALPA 20110816
    Set oCajaIF = New DCajaCtasIF
    If gsOpeCod = gOpeCGAdeudaExtProvisiónMN Or gsOpeCod = gOpeCGAdeudaExtProvisiónME Then
        Call oCajaIF.ActualizarCtaIFCalendario(1, lsPersCod, lsIFTpo, lsCtaIFCod, lnNroCuota)
    ElseIf gsOpeCod = gOpeCGAdeudaExtPagoCuotaMN Or gsOpeCod = gOpeCGAdeudaExtPagoCuotaME Then
        Call oCajaIF.ActualizarCtaIFCalendario(2, lsPersCod, lsIFTpo, lsCtaIFCod, lnNroCuota)
    End If
    Set oCajaIF = Nothing
    
    If oCaja.GrabaExtornoMov(gdFecSis, ldFechaMov, lsMovNroExt, lnMovNro, gsOpeCod, txtMovDesc, lnImporte, lsMovNroExt, lbEliminaMov, lsMovNro, lsAgeCodRef, lsDocNro, lsObjetoCod, gbBitCentral, ldFecReg, lnNumTran) = 0 Then
        Set oCon = Nothing
        If Not lbEliminaMov Then
            ImprimeAsientoContable lsMovNroExt, , , , True, False, Me.txtMovDesc
        End If
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & "Se Grabo el Extorno "
            Set objPista = Nothing
            '*******
            
        fgListaCG.EliminaFila fgListaCG.row
        If MsgBox("Desea realizar otro extorno de Movimiento??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            txtMovDesc = ""
            txtConcepto = ""
            fgListaCG.SetFocus
        Else
            Unload Me
            Exit Sub
        End If
    End If
    cmdExtornar.Enabled = True
End If
Exit Sub
ExtornarErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
    cmdExtornar.Enabled = True
End Sub

Private Sub cmdProcesar_Click()
Dim rs As ADODB.Recordset
Dim lsOperacion As String
Dim oOpe As New DOperacion
Set rs = New ADODB.Recordset
If ValFecha(Me.txtDesde) = False Then Exit Sub
If ValFecha(txthasta) = False Then Exit Sub

If CDate(txtDesde) > CDate(txthasta) Then
    MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
    Exit Sub
End If

lsOperacion = oOpe.GetOperacionRefencia(gsOpeCod)
Select Case gsOpeCod
    Case gOpeMEExtCompraAInst
        Set rs = oCaja.GetMovCompraVentaME(gOpeMECompraAInst, CDate(txtDesde), CDate(txthasta))
    Case gOpeMEExtVentaAInst
        Set rs = oCaja.GetMovCompraVentaME(gOpeMEVentaAInst, CDate(txtDesde), CDate(txthasta))
    Case gOpeMEExtCompraEfect
        Set rs = oCaja.GetMovCompraVentaEfectivo(gOpeMECompraEfect, CDate(txtDesde), CDate(txthasta))
    Case gOpeMEExtVentaEfec
        Set rs = oCaja.GetMovCompraVentaEfectivo(gOpeMEVentaEfec, CDate(txtDesde), CDate(txthasta))
    Case gOpeCGTransfExtBancosMN, gOpeCGTransfExtBancosME
        Set rs = oCaja.GetMovCompraVentaME(lsOperacion, CDate(txtDesde), CDate(txthasta))
    Case gOpeCGTransfExtCMACSBancosMN, gOpeCGTransfExtCMACSBancosME
        Set rs = oCaja.GetMovCompraVentaME(lsOperacion, CDate(txtDesde), CDate(txthasta))
    Case gOpeCGTransfExtBancosCMACSMN, gOpeCGTransfExtBancosCMACSME
        Set rs = oCaja.GetMovCompraVentaME(lsOperacion, CDate(txtDesde), CDate(txthasta))
    Case gOpeCGTransfExtMismoBancoMN, gOpeCGTransfExtMismoBancoME
        Set rs = oCaja.GetMovCompraVentaME(lsOperacion, CDate(txtDesde), CDate(txthasta))
    Case "401435", "402435" 'Extorno Trans. entre agencias Add GITU 03/07/2008
        Set rs = oCaja.GetMovCompraVentaME(lsOperacion, CDate(txtDesde), CDate(txthasta))
    Case gOpeCGOpeBancosConfRetEfecMN, gOpeCGOpeBancosConfRetEfecME
        lsOperacion = IIf(gsOpeCod = gOpeCGOpeBancosConfRetEfecMN, gOpeCGOpeBancosRetEfecMN, gOpeCGOpeBancosRetEfecME)
        Set rs = oCaja.GetRetirosEfectivoCaja(lsOperacion, CDate(txtDesde), CDate(txthasta))
    Case gOpeCGOpeConfApertMN, gOpeCGOpeConfApertME, gOpeCGOpeCMACConfAperMN, gOpeCGOpeCMACConfAperME
        Set rs = oCaja.GetApertSinConf(gsOpeCod, CDate(txtDesde), CDate(txthasta))

    Case gOpeCGAdeudaRegPagareConfMN, gOpeCGAdeudaRegPagareConfMe, _
         gOpeCGAdeudaExtRegistroMN, gOpeCGAdeudaExtRegistroME
        Set rs = oCaja.GetPagareSinConf(gsOpeCod, CDate(txtDesde), CDate(txthasta))
    
    
    'Extornos Regularizacion de Ingreso por Ventanilla
    Case gOpeCGRVentanaIngresoMNExt, gOpeCGRVentanaIngresoMEExt
        Set rs = oCaja.GetCajaRegOpeVentanilla(lsOperacion, txtDesde, txthasta)
    
    'Extornos Autorizacion de Egreso por Ventanilla
    Case gOpeCGRVentanaEgresoMNExt, gOpeCGRVentanaEgresoMEExt
         
    'Extorno de Retiro de Efectivo, No mostrar los retiros confirmados
    Case gOpeCGExtBcoRetEfectivo, gOpeCGExtBcoRetEfectivoME
        Set rs = oCaja.GetCajaBancosOperaciones(lsOperacion, txtDesde, txthasta, True, IIf(Mid(gsOpeCod, 3, 1) = gMonedaExtranjera, gOpeCGOpeBancosConfRetEfecME, gOpeCGOpeBancosConfRetEfecMN))
    
    'Extornos de Dep/Ret de Efectivo de Bancos
    Case gOpeCGExtBcoDepEfectivo, gOpeCGExtBcoDepEfectivoME, _
         gOpeCGExtBcoConfRetEfectivo, gOpeCGExtBcoConfRetEfectivoME, _
         gOpeCGExtBcoGastComision, gOpeCGExtBcoGastComisionME, _
         gOpeCGExtBcoIntCtasAho, gOpeCGExtBcoIntCtasAhoME, _
         gOpeCGOpeCMACDepDivMNExt, gOpeCGOpeCMACDepDivMEExt, _
         gOpeCGOpeCMACRetDivMNExt, gOpeCGOpeCMACRetDivMEExt, _
         gOpeCGOpeCMACRegularizMNExt, gOpeCGOpeCMACRegularizMEExt
        Set rs = oCaja.GetCajaBancosOperaciones(lsOperacion, txtDesde, txthasta)

    'Para Operaciones de Depositos/Retiros diversos tener en cuenta
    'si se referenció Operación de Agencia
    Case gOpeCGExtBcoRetDiv, gOpeCGExtBcoRetDivME, _
         gOpeCGExtBcoDepDiv, gOpeCGExtBcoDepDivME, _
         gOpeCGOpeBancosOtrosDepositosMNExt, gOpeCGOpeBancosOtrosRetirosMNExt, _
         gOpeCGOpeBancosOtrosDepositosMEExt, gOpeCGOpeBancosOtrosRetirosMEExt
         
        Set rs = oCaja.GetCajaBancosOperaciones(lsOperacion, txtDesde, txthasta, , , True)
    
    'Registro de Cheques en Caja General
    Case gOpeCGExtBcoRegCheques, gOpeCGExtBcoRegChequesME
        Set rs = oCaja.GetCajaExtChequesRegSinDeposito(lsOperacion, txtDesde, txthasta)
    'Extorno de Recepcion de Cheques Recibidos en Agencias
    Case gOpeCGExtBcoRecepChqRegAgencias, gOpeCGExtBcoRecepChqRegAgenciasME
        Set rs = oCaja.GetCajaExtChequesRegSinDeposito(lsOperacion, txtDesde, txthasta)

    'Extornos de Ope de Cheques
    Case gOpeCGExtBcoDepCheques, gOpeCGExtBcoDepChequesME
        Set rs = oCaja.GetCajaBancosOperaciones(lsOperacion, txtDesde, txthasta)

         
    'Extorno de Cuentas de Bancos
    Case gOpeCGExtBcoApertCta, gOpeCGExtBcoApertCtaME, _
         gOpeCGExtBcoConfApert, gOpeCGExtBcoConfApertME, _
         gOpeCGExtBcoIntDevengPF, gOpeCGExtBcoIntDevengPFME, _
         gOpeCGExtBcoCapitalizaIntDPF, gOpeCGExtBcoCapitalizaIntDPFME
         
        Set rs = oCaja.GetCajaBancosOperaciones(lsOperacion, txtDesde, txthasta, , , , , gsOpeCod)
    
    Case gOpeCGExtBcoCancelaCtas, gOpeCGExtBcoCancelaCtasME
        Set rs = oCaja.GetCajaBancosOperaciones(lsOperacion, txtDesde, txthasta, , , , "<")
    
    'Extorno de Cuentas de Cmacs
    Case gOpeCGOpeCMACAperCtasMNExt, gOpeCGOpeCMACAperCtasMEExt, _
        gOpeCGOpeCMACConfAperMNExt, gOpeCGOpeCMACConfAperMEExt, _
        gOpeCGOpeCMACIntDevPFMNExt, gOpeCGOpeCMACIntDevPFMEExt, _
        gOpeCGOpeCMACGastosComMNExt, gOpeCGOpeCMACGastosComMEExt, _
        gOpeCGOpeCMACCapIntDevPFMNExt, gOpeCGOpeCMACCapIntDevPFMEExt, _
        gOpeCGOpeCMACInteresAhoMNExt, gOpeCGOpeCMACInteresAhoMEExt
        
        Set rs = oCaja.GetCajaBancosOperaciones(lsOperacion, txtDesde, txthasta, , , , , gsOpeCod)
    
    Case gOpeCGOpeCMACCancelaMNExt, gOpeCGOpeCMACCancelaMEExt
        Set rs = oCaja.GetCajaBancosOperaciones(lsOperacion, txtDesde, txthasta, , , , "<")

    
    
    
    'Extornos Pagarés Adeudados
   ' Case gOpeCGAdeudaExtRegistroMN, gOpeCGAdeudaExtRegistroME
        'Set rs = oCaja.GetPagareExt(lsOperacion, CDate(txtDesde), CDate(txthasta), gEstadoCtaIFActiva)
     
    Case gOpeCGAdeudaExtConfRegiMN, gOpeCGAdeudaExtConfRegiME
        Set rs = oCaja.GetPagareSinConf(lsOperacion, CDate(txtDesde), CDate(txthasta), gEstadoCtaIFActiva)
    
    Case gOpeCGAdeudaExtProvisiónMN, gOpeCGAdeudaExtProvisiónME, _
         gOpeCGAdeudaExtReprogramaMN, gOpeCGAdeudaExtReprogramaME
        
        
        
        Set rs = oAdeud.GetAdeudadosOperaciones(lsOperacion, txtDesde, txthasta)
    
    Case gOpeCGAdeudaExtPagoCuotaMN, gOpeCGAdeudaExtPagoCuotaME
'        Set rs = oAdeud.GetAdeudadosPagoCuota(lsOperacion, txtDesde, txthasta)
        'Set rs = oAdeud.GetUltimosMovimientosExtorno
        Set rs = oAdeud.GetUltimosMovimientosExtorno11(lsOperacion, "", Format(CDate(Me.txtDesde), gsFormatoMovFecha), Format(CDate(txthasta) + 1, gsFormatoMovFecha))
        'Set rs = oAdeud.GetUltimosMovimientosExtorno11(lsOperacon, "", Format(CDate(Me.txtDesde), gsFormatoMovFecha), Format(CDate(txtHasta) + 1, gsFormatoMovFecha))

    
    
    Case OpeCGPagoSunatExt, OpeCGPagoSunatSegSocialExt, OpeCGPagoSunatBoletaPagExt, OpeCGPagoSunatIGVRentaExt
        Set rs = oCaja.GetCajaOperacionesGen(lsOperacion, txtDesde, txthasta)
    
    Case OpeCGCartaFianzaIngExt, OpeCGCartaFianzaIngMEExt
         Set rs = oCaja.GetDatosCartaFianza(oOpe.EmiteOpeCta(IIf(Mid(gsOpeCod, 3, 1) = "1", OpeCGCartaFianzaSal, OpeCGCartaFianzaSalME), "D"), 0, Format(txtDesde, gsFormatoMovFecha), Format(txthasta, gsFormatoMovFecha))

    Case OpeCGCartaFianzaSalExt, OpeCGCartaFianzaSalMEExt
         'Set rs = oCaja.GetDatosCartaFianza(oOpe.EmiteOpeCta(IIf(Mid(gsOpeCod, 3, 1) = "1", OpeCGCartaFianzaRepSalida, OpeCGCartaFianzaRepSalidaME), "D"), 2, Format(txtDesde, gsFormatoMovFecha), Format(txthasta, gsFormatoMovFecha))
         Set rs = oCaja.GetDatosCartaFianza(oOpe.EmiteOpeCta(IIf(Mid(gsOpeCod, 3, 1) = "1", OpeCGCartaFianzaSal, OpeCGCartaFianzaSal), "D"), 2, Format(txtDesde, gsFormatoMovFecha), Format(txthasta, gsFormatoMovFecha))

    Case OpeCGOtrosOpeEfecIngrExt, OpeCGOtrosOpeEfecEgreExt, _
         OpeCGOtrosOpeEfecCambExt, OpeCGOtrosOpeEfecOtroExt, _
         OpeCGOtrosOpeEfecIngrMEExt, OpeCGOtrosOpeEfecEgreMEExt, _
         OpeCGOtrosOpeEfecCambMEExt, OpeCGOtrosOpeEfecOtroMEExt
        Set rs = oCaja.GetCajaOperacionesGen(lsOperacion, txtDesde, txthasta)
    
    Case gOpeCGExtProvPago, gOpeCGExtProvPagoEfectivo, gOpeCGExtProvPagoTransfer, _
         gOpeCGExtProvPagoOPago, gOpeCGExtProvPagoCheque, _
         gOpeCGExtProvPagoAbono, gOpeCGExtProvPagoRechazo, _
         gOpeCGExtProvPagoME, gOpeCGExtProvPagoEfectivoME, gOpeCGExtProvPagoTransferME, _
         gOpeCGExtProvPagoOPagoME, gOpeCGExtProvPagoChequeME, _
         gOpeCGExtProvPagoAbonoME, gOpeCGExtProvPagoRechazoME
        Set rs = oCaja.GetCajaOperacionesGen(lsOperacion, txtDesde, txthasta, , , 0)
    
'JEOM
    Case gOpeCGExtProvPagoSUNAT
        Set rs = oCaja.GetCajaOperacionesGen(lsOperacion, txtDesde, txthasta, , , 5)
'FIN

    Case gOpeCGExtProvEntrCheques, gOpeCGExtProvEntrChequesME, gOpeCGExtProvEntrOPago, gOpeCGExtProvEntrOPagoME
        Set rs = oCaja.GetCajaGenConfirmacion(lsOperacion, txtDesde, txthasta, , , 0)
    
    'Confirmación de Retiro de Efectivo con CMACS
    '***Modificado por ELRO el 20110923, según Acta 263-2011/TI-D
    Case gOpeCGOpeCMACConRetEfeMN, gOpeCGOpeCMACConRetEfeME
    
         lsOperacion = IIf(gsOpeCod = gOpeCGOpeCMACConRetEfeMN, gOpeCGOpeCMACRetEfeMN, gOpeCGOpeCMACRetEfeME)
         Set rs = oCaja.GetRetirosEfectivoCaja(lsOperacion, CDate(txtDesde), CDate(txthasta))
    '***Fin Modificado por ELRO**********************************
    
    'Extorno de Deposito, Retiro y Confirmación de Efectivo con CMACS
    '***Modificado por ELRO el 20110923, según Acta 269-2011/TI-D
    Case gOpeCGOpeCMACExtDepEfeMN, gOpeCGOpeCMACExtDepEfeME, _
         gOpeCGOpeCMACExtRetEfeMN, gOpeCGOpeCMACExtRetEfeME, _
         gOpeCGOpeCMACExtConRetEfeMN, gOpeCGOpeCMACExtConRetEfeME
    
     Set rs = oCaja.GetCajaBancosOperaciones(lsOperacion, txtDesde, txthasta)
    '***Fin Modificado por ELRO**********************************
    
    Case gOpeCGExtProvDevSUNAT 'PASIERS1242014
        Set rs = oCaja.GetDatosPagoDevolucionProveedorSunatxExtorno(txtDesde, txthasta)
    'END PASI
    Case OpeCGOtrosOpeRetPagSeguroDesgravamenMNExt, OpeCGOtrosOpeRetPagSeguroIncendioMNExt, _
            OpeCGOtrosOpeRetPagSeguroDesgravamenMEExt, OpeCGOtrosOpeRetPagSeguroIncendioMEExt 'PASIERS1362014
    Dim lsOpeCod As String
    Dim oNSeg As NSeguros
    Set oNSeg = New NSeguros
    Select Case gsOpeCod
        Case OpeCGOtrosOpeRetPagSeguroDesgravamenMNExt, OpeCGOtrosOpeRetPagSeguroIncendioMNExt
            lsOpeCod = IIf(gsOpeCod = OpeCGOtrosOpeRetPagSeguroDesgravamenMNExt, OpeCGOtrosOpeRetPagSeguroDesgravamenMN, OpeCGOtrosOpeRetPagSeguroIncendioMN)
        Case OpeCGOtrosOpeRetPagSeguroDesgravamenMEExt, OpeCGOtrosOpeRetPagSeguroIncendioMEExt
            lsOpeCod = IIf(gsOpeCod = OpeCGOtrosOpeRetPagSeguroDesgravamenMEExt, OpeCGOtrosOpeRetPagSeguroDesgravamenME, OpeCGOtrosOpeRetPagSeguroIncendioME)
    End Select
        Set rs = oNSeg.GetMovRetPagoxExtorno(CDate(txtDesde.Text), CDate(txthasta.Text), lsOpeCod)
    'END PASI
    Case Else
        Set rs = oCaja.GetCajaOperacionesGen(lsOperacion, txtDesde, txthasta)
    
End Select
fgListaCG.Clear
fgListaCG.FormaCabecera
fgListaCG.Rows = 2
If rs Is Nothing Then
    MsgBox "Operación no implementada totalmente", vbInformation, "¡Aviso!"
Else
    If rs.State = adStateClosed Then
        MsgBox "Operación no implementada totalmente", vbInformation, "¡Aviso!"
    Else
        
    If Not rs.EOF And Not rs.BOF Then
        Set fgListaCG.Recordset = rs
        fgListaCG.SetFocus
    Else
        MsgBox "Datos no encontrados para proceso seleccionado", vbInformation, "Aviso"
    End If
    End If
End If
fgListaCG_RowColChange
RSClose rs

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgListaCG_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMovDesc.SetFocus
End If
End Sub

Private Sub fgListaCG_RowColChange()
'On Error Resume Next
Dim lsTipoCtaIf As String 'EJVG20120802
If fgListaCG.TextMatrix(1, 0) = "" Then
    txtMovDesc = ""
    txtConcepto = ""
    Me.txtCuenta = ""
    Me.txtSubCuenta = ""
    Exit Sub
End If
lsTipoCtaIf = Mid(fgListaCG.TextMatrix(fgListaCG.row, 8), 1, 2) 'EJVG20120802
Select Case gsOpeCod
    Case gOpeMEExtCompraAInst, gOpeMEExtVentaAInst
        txtConcepto = fgListaCG.TextMatrix(fgListaCG.row, 8)
    Case gOpeMEExtCompraEfect, gOpeMEExtVentaEfec
        txtConcepto = fgListaCG.TextMatrix(fgListaCG.row, 2)
    Case gOpeCGRVentanaIngresoMNExt, gOpeCGRVentanaIngresoMEExt
        txtConcepto = fgListaCG.TextMatrix(fgListaCG.row, 5)
    Case gOpeCGTransfExtBancosMN, gOpeCGTransfExtCMACSBancosMN, gOpeCGTransfExtBancosCMACSMN, gOpeCGTransfExtMismoBancoMN, "401435", _
        gOpeCGTransfExtBancosME, gOpeCGTransfExtCMACSBancosME, gOpeCGTransfExtBancosCMACSME, gOpeCGTransfExtMismoBancoME, "402435", _
        gOpeCGOpeBancosConfRetEfecMN, gOpeCGOpeBancosConfRetEfecME, _
        "401607", "402607" 'gOpeCGOpeCMACConRetEfeMN, gOpeCGOpeCMACConRetEfeME
        txtConcepto = fgListaCG.TextMatrix(fgListaCG.row, 8)
    Case gOpeCGOpeConfApertMN, gOpeCGOpeConfApertME, gOpeCGOpeCMACConfAperMN, gOpeCGOpeCMACConfAperME, gOpeCGAdeudaRegPagareConfMN, gOpeCGAdeudaRegPagareConfMe
        Dim lsCuenta As String, lsSubCuenta As String
        txtConcepto = fgListaCG.TextMatrix(fgListaCG.row, 9)
        txtCuenta = "": txtSubCuenta = ""
        txtCuenta.Enabled = True: txtSubCuenta.Enabled = True
        If Mid(fgListaCG.TextMatrix(fgListaCG.row, 8), 1, 2) = "05" Then
        'edpyme
           oCaja.GeneraCuentaAdeudos lsCuenta, lsSubCuenta, fgListaCG.TextMatrix(fgListaCG.row, 11), fgListaCG.TextMatrix(fgListaCG.row, 10), "", gsOpeCod, IIf(gsOpeCod = gOpeCGAdeudaRegPagareConfMN Or gsOpeCod = gOpeCGAdeudaRegPagareConfMe, "H", "D")
        Else
            'oCaja.GeneraCuentaBancos lsCuenta, lsSubCuenta, fgListaCG.TextMatrix(fgListaCG.Row, 11), fgListaCG.TextMatrix(fgListaCG.Row, 10), fgListaCG.TextMatrix(fgListaCG.Row, 8), gsOpeCod, IIf(gsOpeCod = gOpeCGAdeudaRegPagareConfMN Or gsOpeCod = gOpeCGAdeudaRegPagareConfMe, "H", "D")
            'EJVG20120802 ***
            If lsTipoCtaIf = "03" Or lsTipoCtaIf = "04" Then
                oCaja.GeneraCuentaBancos2 lsCuenta, lsSubCuenta, fgListaCG.TextMatrix(fgListaCG.row, 11), fgListaCG.TextMatrix(fgListaCG.row, 10), fgListaCG.TextMatrix(fgListaCG.row, 8), gsOpeCod, IIf(gsOpeCod = gOpeCGAdeudaRegPagareConfMN Or gsOpeCod = gOpeCGAdeudaRegPagareConfMe, "H", "D")
            Else
                oCaja.GeneraCuentaBancos lsCuenta, lsSubCuenta, fgListaCG.TextMatrix(fgListaCG.row, 11), fgListaCG.TextMatrix(fgListaCG.row, 10), fgListaCG.TextMatrix(fgListaCG.row, 8), gsOpeCod, IIf(gsOpeCod = gOpeCGAdeudaRegPagareConfMN Or gsOpeCod = gOpeCGAdeudaRegPagareConfMe, "H", "D")
            End If
            'END EJVG *******
        End If
        If lsCuenta <> "" Then
            txtCuenta = lsCuenta
            'txtCuenta.Enabled = False
            If lsSubCuenta <> "" Then
                txtSubCuenta = lsSubCuenta
            End If
        Else
            If lsSubCuenta <> "" Then
                If Me.fgListaCG.TextMatrix(fgListaCG.row, 12) = 1 Then
                    txtSubCuenta = Mid(lsSubCuenta, 1, Len(lsSubCuenta) - 1) & "4"
                Else
                    txtSubCuenta = lsSubCuenta
                End If
            End If
        End If
        If Err Then
            MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
            Err.Clear
        End If
        
    Case gOpeCGAdeudaExtRegistroMN, gOpeCGAdeudaExtRegistroME, _
         gOpeCGAdeudaExtConfRegiMN, gOpeCGAdeudaExtConfRegiME
        txtConcepto = fgListaCG.TextMatrix(fgListaCG.row, 9)
         
    Case gOpeCGAdeudaExtProvisiónMN, gOpeCGAdeudaExtProvisiónMN, _
         gOpeCGAdeudaExtPagoCuotaMN, gOpeCGAdeudaExtPagoCuotaME, _
         gOpeCGAdeudaExtReprogramaMN, gOpeCGAdeudaExtReprogramaME
        txtConcepto = fgListaCG.TextMatrix(fgListaCG.row, 5)
    
    Case OpeCGCartaFianzaIngExt, OpeCGCartaFianzaSalExt, _
         OpeCGCartaFianzaIngMEExt, OpeCGCartaFianzaSalMEExt
        txtConcepto = fgListaCG.TextMatrix(fgListaCG.row, 2)
        
    Case Else
        txtConcepto = fgListaCG.TextMatrix(fgListaCG.row, 7)
End Select
txtMovDesc = txtConcepto

End Sub

Private Sub Form_Load()
Set oCaja = New nCajaGeneral
Set oAdeud = New NCajaAdeudados
Set objPista = New COMManejador.Pista

'***Modificado por ELRO el 20120103, según Acta 311-2011/TI-D
Dim oNConstSistemas As New NConstSistemas
ldFecCie = CDate(oNConstSistemas.LeeConstSistema(gConstSistCierreMensualCont))
Set oNConstSistemas = Nothing
'***Fin Modificado por ELRO**********************************

CentraForm Me
Me.Caption = gsOpeDesc
txtDesde = gdFecSis
txthasta = gdFecSis
cmdConfAper.Visible = False
cmdConfRetiro.Visible = False
cmdExtornar.Visible = False
FraFechaMov.Visible = False
txtFechaMov = gdFecSis
Select Case gsOpeCod
    Case gOpeMEExtCompraAInst, gOpeMEExtVentaAInst
        fgListaCG.EncabezadosAnchos = "350-900-2000-1500-2000-1500-1200-0-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Institución Origen -Cuenta-Institución Destino-Cuenta-Importe-cMovNro-Concepto"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-L-L-R-L"
        fgListaCG.FormatosEdit = "0-0-0-0-0-0-2-0-0"
        cmdExtornar.Visible = True

    Case gOpeMEExtCompraEfect, gOpeMEExtVentaEfec
        fgListaCG.EncabezadosNombres = "N°-Fecha-Concepto-Monto Dolares - Monto Soles- Cambio-cMovNro"
        fgListaCG.EncabezadosAnchos = "350-900-4500-1400-1400-1000-0"
        fgListaCG.EncabezadosAlineacion = "C-C-L-R-R-R-C"
        fgListaCG.FormatosEdit = "0-0-0-2-2-2-0"
        cmdExtornar.Visible = True
    Case gOpeCGTransfExtBancosMN, gOpeCGTransfExtCMACSBancosMN, gOpeCGTransfExtBancosCMACSMN, gOpeCGTransfExtMismoBancoMN, "401435", _
        gOpeCGTransfExtBancosME, gOpeCGTransfExtCMACSBancosME, gOpeCGTransfExtBancosCMACSME, gOpeCGTransfExtMismoBancoME, "402435"
        fgListaCG.EncabezadosAnchos = "350-900-2000-1500-2000-1500-1200-0-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Institución Origen -Cuenta-Institución Destino-Cuenta-Importe-cMovNro-Concepto"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-L-L-R-L"
        fgListaCG.FormatosEdit = "0-0-0-0-0-0-2-0-0"
        cmdExtornar.Visible = True
    Case gOpeCGOpeBancosConfRetEfecMN, gOpeCGOpeBancosConfRetEfecME
        fgListaCG.EncabezadosAnchos = "350-600-1000-1000-2000-3300-1200-0-0"
        fgListaCG.EncabezadosNombres = "N°-Tipo-Número-Voucher-Banco-Cuenta-Importe-nMovNro-Concepto"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-L-L-R-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-0-0-2-0-0"
        lbltitulo = ""
        cmdConfRetiro.Visible = True
        txtDesde = DateAdd("d", (Day(gdFecSis) - 1) * -1, gdFecSis)
        txthasta = gdFecSis
    Case gOpeCGOpeConfApertMN, gOpeCGOpeConfApertME, gOpeCGOpeCMACConfAperMN, _
         gOpeCGOpeCMACConfAperME, gOpeCGAdeudaRegPagareConfMN, gOpeCGAdeudaRegPagareConfMe
        fgListaCG.EncabezadosAnchos = "350-900-2500-3500-900-1200-0-0-0-0-0-0-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Banco-Cuenta-Apertura-Importe-nMovNro-cOpeCod-cCtaIfCod-Concepto-cIfTpo-cPersCod-Tipo"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-C-R-L-L-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-0-2-0-0-0-0-0-0"
        lbltitulo = "APERTURAS..."
        If gsOpeCod = gOpeCGAdeudaRegPagareConfMN Or gsOpeCod = gOpeCGAdeudaRegPagareConfMe Then
            lbltitulo = "ADEUDADOS..."
        End If
        cmdConfAper.Visible = True
        txtDesde = DateAdd("d", (Day(gdFecSis) - 1) * -1, gdFecSis)
        txthasta = gdFecSis
        FraFechaMov.Visible = True
        txtMovDesc.Height = 385
        lblSubCuenta.Visible = True
        txtSubCuenta.Visible = True
        txtCuenta.Visible = True
        lblCuenta.Visible = True
        Dim oOpe As New DOperacion
        If gsOpeCod = gOpeCGAdeudaRegPagareConfMN Or gsOpeCod = gOpeCGAdeudaRegPagareConfMe Then
           txtCuenta.rs = oOpe.CargaOpeCtaArbol(gsOpeCod, "H", "0")
        Else
           txtCuenta.rs = oOpe.CargaOpeCtaArbol(gsOpeCod, "D", "0")
        End If
        txtCuenta.psRaiz = "Cuentas Contables"
        Set oOpe = Nothing
    
    'Operaciones con Ventanilla
    Case gOpeCGRVentanaIngresoMNExt, gOpeCGRVentanaEgresoMNExt, _
         gOpeCGRVentanaIngresoMEExt, gOpeCGRVentanaEgresoMEExt
        fgListaCG.EncabezadosAnchos = "350-1000-3500-1400-1200-0-0-2200-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Persona -Nro Operación-Importe-Concepto-cPersCod-cMovNro"
        fgListaCG.EncabezadosAlineacion = "C-C-L-C-R-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-2-0-0-0-0"
        cmdExtornar.Visible = True
    
    Case gOpeCGExtBcoRecepChqRegAgencias, gOpeCGExtBcoRecepChqRegAgenciasME, _
         gOpeCGExtBcoRegCheques, gOpeCGExtBcoRegChequesME, _
        fgListaCG.EncabezadosAnchos = "350-850-2600-1700-400-1200-1200-0-0-0-0-2050-0-0-0-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Institución Financiera-Cuenta-Doc-Nro.Doc-Importe-Concepto-cPersCod-cIFTpo-cCtaIFCod-cMovNro-nMovNro-nDocTpo-cAreaCod-cAgeCod"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-C-C-R-L-L-L-L-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-0-0-2-0-0-0-0-0-0-0-0"
        cmdExtornar.Visible = True
         
    Case gOpeCGExtBcoDepCheques, gOpeCGExtBcoDepChequesME
        fgListaCG.EncabezadosAnchos = "350-850-2600-1700-0-0-1200-0-0-0-0-2050-0-0-0-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Institución Financiera-Cuenta-Doc-Nro.Doc-Importe-Concepto-cPersCod-cIFTpo-cCtaIFCod-cMovNro-nMovNro-nDocTpo-cAreaCod-cAgeCod"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-C-C-R-L-L-L-L-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-0-0-2-0-0-0-0-0-0-0-0"
        cmdExtornar.Visible = True
         
    'Extornos de Bancos
    Case gOpeCGExtBcoDepEfectivo, gOpeCGExtBcoRetEfectivo, _
         gOpeCGExtBcoConfRetEfectivo, gOpeCGExtBcoDepDiv, _
         gOpeCGExtBcoRetDiv, gOpeCGExtBcoApertCta, _
         gOpeCGExtBcoConfApert, gOpeCGExtBcoIntDevengPF, _
         gOpeCGExtBcoGastComision, gOpeCGExtBcoCapitalizaIntDPF, _
         gOpeCGExtBcoCancelaCtas, gOpeCGExtBcoIntCtasAho, _
         gOpeCGExtBcoDepEfectivoME, gOpeCGExtBcoRetEfectivoME, _
         gOpeCGExtBcoConfRetEfectivoME, gOpeCGExtBcoDepDivME, _
         gOpeCGExtBcoRetDivME, _
         gOpeCGExtBcoApertCtaME, gOpeCGExtBcoConfApertME, _
         gOpeCGExtBcoIntDevengPFME, gOpeCGExtBcoGastComisionME, _
         gOpeCGExtBcoCapitalizaIntDPFME, gOpeCGExtBcoCancelaCtasME, _
         gOpeCGExtBcoIntCtasAhoME, _
         gOpeCGOpeBancosOtrosDepositosMNExt, gOpeCGOpeBancosOtrosRetirosMNExt, _
         gOpeCGOpeBancosOtrosDepositosMEExt, gOpeCGOpeBancosOtrosRetirosMEExt
        
        fgListaCG.EncabezadosAnchos = "350-850-2600-1700-400-1200-1200-0-0-0-0-2050-0-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Institución Financiera-Cuenta-Doc-Nro.Doc-Importe-Concepto-cPersCod-cIFTpo-cCtaIFCod-cMovNro-nMovNro-nDocTpo"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-C-C-R-L-L-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-0-0-2-0-0-0-0-0-0"
        cmdExtornar.Visible = True
    
    'Extornos Adeudados
    Case gOpeCGAdeudaExtRegistroMN, gOpeCGAdeudaExtRegistroME, _
         gOpeCGAdeudaExtConfRegiMN, gOpeCGAdeudaExtConfRegiME
        fgListaCG.EncabezadosAnchos = "350-900-2500-3500-900-1200-0-0-0-0-0-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Banco-Cuenta-Apertura-Importe-nMovNro-cOpeCod-cCtaIfCod-Concepto-cIfTpo-cPersCod"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-C-R-L-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-0-2-0-0-0-0-0"
        cmdExtornar.Visible = True
        
    Case gOpeCGAdeudaExtProvisiónMN, gOpeCGAdeudaExtProvisiónMN, _
         gOpeCGAdeudaExtReprogramaMN, gOpeCGAdeudaExtReprogramaME
        
        fgListaCG.EncabezadosAnchos = "350-850-2600-1700-1200-0-0-0-0-2050-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Institución Financiera-Cuenta-Importe-Concepto-cPersCod-cIFTpo-cCtaIFCod-cMovNro-nMovNro"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-R-L-L-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-2-0-0-0-0-0-0"
        cmdExtornar.Visible = True

    Case gOpeCGAdeudaExtPagoCuotaMN, gOpeCGAdeudaExtPagoCuotaME
        fgListaCG.EncabezadosAnchos = "350-850-2600-1700-1200-0-0-0-0-2050-0-2500"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Institución Financiera-Cuenta-Capital-Concepto-cPersCod-cIFTpo-cCtaIFCod-cMovNro-nMovNro-Entidad Pago"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-R-L-L-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-2-0-0-0-0-0-0"
        cmdExtornar.Visible = True

    Case OpeCGCartaFianzaIngExt, OpeCGCartaFianzaSalExt, _
         OpeCGCartaFianzaIngMEExt, OpeCGCartaFianzaSalMEExt
        fgListaCG.EncabezadosAnchos = "350-1300-0-2300-2300-1200-0-1100-0-1100-0-0-0"
        fgListaCG.EncabezadosNombres = "N°-Documento-Concepto-Institución Financiera-Proveedor-Importe-Fecha-Fec.Venc.-cIFPersCod-Fecha Ingr.-cMovNro-nMovNro-cMovNroSalida"
        fgListaCG.EncabezadosAlineacion = "C-L-L-L-L-R-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-0-2-0-0-0-0"
        cmdExtornar.Visible = True

    Case OpeCGOtrosOpeEfecIngrExt, OpeCGOtrosOpeEfecEgreExt, _
         OpeCGOtrosOpeEfecCambExt, OpeCGOtrosOpeEfecOtroExt, _
         OpeCGOtrosOpeEfecIngrMEExt, OpeCGOtrosOpeEfecEgreMEExt, _
         OpeCGOtrosOpeEfecCambMEExt, OpeCGOtrosOpeEfecOtroMEExt
    
        fgListaCG.EncabezadosAnchos = "350-1000-700-1300-1100-0-3100-0-2000-0-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Doc-Nro.Doc-Importe-cPersCod-Persona-Concepto-cMovNro-nMovNro-nDocTpo"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-R-C-L-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-2-0-0-0-0-0-0"
        cmdExtornar.Visible = True

    Case gOpeCGOpeCMACDepDivMNExt, gOpeCGOpeCMACDepDivMEExt, _
        gOpeCGOpeCMACRetDivMNExt, gOpeCGOpeCMACRetDivMEExt, _
        gOpeCGOpeCMACRegularizMNExt, gOpeCGOpeCMACRegularizMEExt, _
        gOpeCGOpeCMACAperCtasMNExt, gOpeCGOpeCMACAperCtasMEExt, _
        gOpeCGOpeCMACConfAperMNExt, gOpeCGOpeCMACConfAperMEExt, _
        gOpeCGOpeCMACIntDevPFMNExt, gOpeCGOpeCMACIntDevPFMEExt, _
        gOpeCGOpeCMACGastosComMNExt, gOpeCGOpeCMACGastosComMEExt, _
        gOpeCGOpeCMACCapIntDevPFMNExt, gOpeCGOpeCMACCapIntDevPFMEExt, _
        gOpeCGOpeCMACCancelaMNExt, gOpeCGOpeCMACCancelaMEExt, _
        gOpeCGOpeCMACInteresAhoMNExt, gOpeCGOpeCMACInteresAhoMEExt
        
        fgListaCG.EncabezadosAnchos = "350-850-2600-1700-400-1200-1200-0-0-0-0-2050-0-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Institución Financiera-Cuenta-Doc-Nro.Doc-Importe-Concepto-cPersCod-cIFTpo-cCtaIFCod-cMovNro-nMovNro-nDocTpo"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-C-C-R-L-L-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-0-0-2-0-0-0-0-0-0"
        cmdExtornar.Visible = True
    '***Modificado por ELRO el 20110923, según Acta 263-2011/TI-D
    Case gOpeCGOpeCMACConRetEfeMN, gOpeCGOpeCMACConRetEfeME
        fgListaCG.EncabezadosAnchos = "350-600-1000-1000-2000-3300-1200-0-0"
        fgListaCG.EncabezadosNombres = "N°-Tipo-Número-Voucher-Banco-Cuenta-Importe-nMovNro-Concepto"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-L-L-R-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-0-0-2-0-0"
        lbltitulo = ""
        cmdConfRetiro.Visible = True
        txtDesde = DateAdd("d", (Day(gdFecSis) - 1) * -1, gdFecSis)
        txthasta = gdFecSis
        chkConGenAsiento.Visible = False
    '***Fin Modificado por ELRO**********************************
    '***Modificado por ELRO el 20110930, según Acta 269-2011/TI-D
    Case gOpeCGOpeCMACExtDepEfeMN, gOpeCGOpeCMACExtDepEfeME, _
         gOpeCGOpeCMACExtRetEfeMN, gOpeCGOpeCMACExtRetEfeME, _
         gOpeCGOpeCMACExtConRetEfeMN, gOpeCGOpeCMACExtConRetEfeME
         
         fgListaCG.EncabezadosAnchos = "350-850-2600-1700-400-1200-1200-0-0-0-0-2050-0-0"
         fgListaCG.EncabezadosNombres = "N°-Fecha-Institución Financiera-Cuenta-Doc-Nro.Doc-Importe-Concepto-cPersCod-cIFTpo-cCtaIFCod-cMovNro-nMovNro-nDocTpo"
         fgListaCG.EncabezadosAlineacion = "C-C-L-L-C-C-R-L-L-L-L-L-L"
         fgListaCG.FormatosEdit = "0-0-0-0-0-0-2-0-0-0-0-0-0"
         cmdExtornar.Visible = True
    '***Fin Modificado por ELRO**********************************
    Case gOpeCGExtProvDevSUNAT 'PASIERS1242014
        fgListaCG.EncabezadosAnchos = "350-1000-700-1300-1200-0-3100-4000-2000-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Doc-Nro.Doc-Importe-cPersCod-Persona-Concepto-cMovNro-nMovNro"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-R-C-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-2-0-0-0-0-0"
        cmdExtornar.Visible = True
        'END PASI
    Case OpeCGOtrosOpeRetPagSeguroDesgravamenMNExt, OpeCGOtrosOpeRetPagSeguroIncendioMNExt, _
            OpeCGOtrosOpeRetPagSeguroDesgravamenMEExt, OpeCGOtrosOpeRetPagSeguroIncendioMEExt 'PASIERS1362014
        fgListaCG.EncabezadosAnchos = "350-1000-3000-1800-700-1300-1200-4000-2000-0"
        fgListaCG.EncabezadosNombres = "Nº-Fecha-Institución Financiera-Cuenta-Doc-Nro.Doc-Importe-Concepto-cMovNro-nMovNro"
        fgListaCG.FormatosEdit = "0-0-0-0-0-0-0-0-0-0"
        cmdExtornar.Visible = True
        'END PASI
    Case Else
        fgListaCG.EncabezadosAnchos = "350-1000-700-1300-1100-0-3100-0-2000-0-0"
        fgListaCG.EncabezadosNombres = "N°-Fecha-Doc-Nro.Doc-Importe-cPersCod-Persona-Concepto-cMovNro-nMovNro-nDocTpo"
        fgListaCG.EncabezadosAlineacion = "C-C-L-L-R-C-L-L-L-L-L"
        fgListaCG.FormatosEdit = "0-0-0-0-2-0-0-0-0-0-0"
        cmdExtornar.Visible = True

End Select
End Sub
Private Sub txtCuenta_EmiteDatos()
If txtCuenta.psDescripcion <> "" Then
    If txtSubCuenta.Visible Then
        txtSubCuenta.SetFocus
    End If
End If
End Sub

Private Sub txtDesde_GotFocus()
fEnfoque txtDesde
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtDesde) = False Then Exit Sub
    If CDate(txtDesde) > CDate(txthasta) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
    txthasta.SetFocus
End If
End Sub

Private Sub txtHasta_GotFocus()
fEnfoque txthasta
End Sub

Private Sub txthasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txthasta) = False Then Exit Sub
    If CDate(txtDesde) > CDate(txthasta) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
    cmdProcesar.SetFocus
End If
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdExtornar.Visible And cmdExtornar.Enabled Then cmdExtornar.SetFocus
    If cmdConfRetiro.Visible Then cmdConfRetiro.SetFocus
    If cmdConfAper.Visible Then
        If txtCuenta.Enabled Then
            txtCuenta.SetFocus
        ElseIf txtSubCuenta.Enabled Then
            txtSubCuenta.SetFocus
        Else
            cmdConfAper.SetFocus
        End If
    End If
End If
End Sub

Private Sub txtSubCuenta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii, False)
If KeyAscii = 13 Then
    cmdConfAper.SetFocus
End If
End Sub

Private Function Validar() As Boolean
    Validar = True
    If fgListaCG.TextMatrix(1, 0) = "" Then
       MsgBox "No existen Operaciones para confirmar", vbInformation, "¡Aviso!"
       Validar = False
       Exit Function
    End If
    
    If Len(Trim(txtMovDesc)) = 0 Then
        MsgBox "Descripción de Movimiento no ingresado", vbInformation, "Aviso"
        txtMovDesc.SetFocus
        Validar = False
        Exit Function
    End If
End Function

Private Sub ConfirmaOtros(pnImporte As Currency, psPersCod As String, psTipoIF As String, psCtaIFCod As String, pnMovRef As Double)
    Dim lsCtaHaber  As String
    Dim lsSubCta    As String
    Dim lsMovNro    As String
    Dim oCont       As NContFunciones
    Dim oOpe        As DOperacion
    Dim lnTasaVac   As Double
    Dim lsMonedaPago As String
    'On Error GoTo ConfirmaAperErr
    
    Set oOpe = New DOperacion
    Set oCont = New NContFunciones

    lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
    
    If lsCtaHaber = "" Then
        MsgBox "Cuenta de Pendiente no esta definida. Consultar con Sistemas", vbInformation, "¡Aviso!"
        Exit Sub
    End If
    
    If txtCuenta = "" Then
        MsgBox "Aún no se define Cuenta Contable que genera Cuenta de Institución Financiera", vbInformation, "¡AViso!"
        Exit Sub
    End If
    If txtSubCuenta = "" Then
        If MsgBox("Aún no se define SubCuenta que genera Cuenta de Institución Financiera. ¿Desea Continuar? ", vbQuestion + vbYesNo, "¡Advertencia!") = vbNo Then
            Exit Sub
        End If
    End If

    Dim oCta As DCtaCont
    Set oCta = New DCtaCont
    oCta.ExisteCuenta txtCuenta & txtSubCuenta, True
    If MsgBox("Desea Confirmar la Apertura de la Cuenta Seleccionada?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        
        lsMovNro = oCont.GeneraMovNro(txtFechaMov, gsCodAge, gsCodUser)
'        If oCaja.GrabaConfApertura(lsMovNro, gsOpeCod, txtMovDesc, _
'                                    txtCuenta, txtSubCuenta, lsCtaHaber, pnImporte, psPersCod, _
'                                    psTipoIF, psCtaIFCod, pnMovRef, txtFechaMov, lnTasaVac) = 0 Then
'            ImprimeAsientoContable lsMovNro
'            If MsgBox("Desea Realizar otra Confirmación ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
'                fgListaCG.EliminaFila fgListaCG.Row
'                txtMovDesc = ""
'            Else
'                Unload Me
'            End If
'        End If
        'EJVG20120802 *** DPF Y OverNight
        If Left(psCtaIFCod, 2) = "03" Or Left(psCtaIFCod, 2) = "04" Then
            If oCaja.GrabaConfApertura2(lsMovNro, gsOpeCod, txtMovDesc, _
                                        txtCuenta, txtSubCuenta, lsCtaHaber, pnImporte, psPersCod, _
                                        psTipoIF, psCtaIFCod, pnMovRef, txtFechaMov, lnTasaVac) = 0 Then
                ImprimeAsientoContable lsMovNro
                If MsgBox("Desea Realizar otra Confirmación ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                    fgListaCG.EliminaFila fgListaCG.row
                    txtMovDesc = ""
                    fgListaCG_RowColChange
                Else
                    Unload Me
                End If
            End If
        Else
            If oCaja.GrabaConfApertura(lsMovNro, gsOpeCod, txtMovDesc, _
                                        txtCuenta, txtSubCuenta, lsCtaHaber, pnImporte, psPersCod, _
                                        psTipoIF, psCtaIFCod, pnMovRef, txtFechaMov, lnTasaVac) = 0 Then
                ImprimeAsientoContable lsMovNro
                If MsgBox("Desea Realizar otra Confirmación ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                    fgListaCG.EliminaFila fgListaCG.row
                    txtMovDesc = ""
                    fgListaCG_RowColChange
                Else
                    Unload Me
                End If
            End If
        End If
        'END EJVG *******
    End If
End Sub

Private Sub ConfirmaAdeudados(pnImporte As Currency, psPersCod As String, psTipoIF As String, psCtaIFCod As String, pnMovRef As Double)
Dim oCaja           As New nCajaGeneral
'Dim oAdeudado       As New DACGAdeudados
Dim oCon            As New NContFunciones
Dim oAdeu           As New NCajaAdeudados
Dim oIndi           As New DCaja_Adeudados
Dim oCta            As New DCtaCont
Dim lsCuenta        As String
Dim lsSubCuenta     As String
Dim lnFilaActual    As Long
Dim dFechaAper      As Date

Dim lsCtaHaber      As String
Dim lsMovNro        As String
Dim i               As Long

'CP LP
Dim lsCuentaCorto   As String
Dim lsCuentaLargo   As String

'Indice Vac
Dim lnTasaVac       As Double
Dim lsMonedaPago    As String
Dim lsMoneda        As String

Dim nMontoPrestado  As Double
Dim nMontoPrestadoReal As Double
 
On Error GoTo ConfirmarError_

       
    lsCtaHaber = oAdeu.GetOpeCta(gsOpeCod, "D")
    
    lnFilaActual = Me.fgListaCG.row 'lstCabecera.SelectedItem.Index
    dFechaAper = Me.fgListaCG.TextMatrix(lnFilaActual, 1) 'CDate(lstCabecera.ListItems(lnFilaActual).ListSubItems(4).Text)
    nMontoPrestado = CCur(Me.fgListaCG.TextMatrix(lnFilaActual, 13)) 'Val(lstCabecera.ListItems(lnFilaActual).ListSubItems(13).Text)
    nMontoPrestadoReal = CCur(Me.fgListaCG.TextMatrix(lnFilaActual, 14)) 'Val(lstCabecera.ListItems(lnFilaActual).ListSubItems(14).Text)
    lsMonedaPago = Me.fgListaCG.TextMatrix(lnFilaActual, 15) 'Trim(lstCabecera.ListItems(lnFilaActual).ListSubItems(15).Text)
    lsMoneda = Mid(psCtaIFCod, 3, 1)
    
    'oCaja.GeneraCuentaBancos lsCuenta, lsSubCuenta, psPersCod, psTipoIF, psCtaIFCod, gsOpeCod, "H"

    'Verifico si existe cuenta contable
    'If oAdeu.GetExisteCuentaContable(lsCuenta, lsSubCuenta) = True Then
                 
        'VAC
        lnTasaVac = 0
        If lsMoneda = "1" And lsMonedaPago = "2" Then
           lnTasaVac = oIndi.CargaIndiceVAC(dFechaAper)
        End If
    
        'CP LP optener las cuentas de configuracion analoga
        'lsCuentaCorto = oAdeu.GetOpeCta(gsOpeCod, "D", "1", psPersCod, psTipoIF)
        'lsCuentaLargo = oAdeu.GetOpeCta(gsOpeCod, "H", "1", psPersCod, psTipoIF)
        lsCuenta = Me.txtCuenta.Text
        lsSubCuenta = Me.txtSubCuenta.Text
        If Left(Trim(Me.txtCuenta.Text), 2) = "26" Then
            lsCuentaLargo = Me.txtCuenta.Text & Me.txtSubCuenta.Text
            If psPersCod = "1093300012530" Then 'PARA CLIENTES MI VIVIENDA
                lsCuentaCorto = "24" & Mid(lsCuentaLargo, 3, 1) & "601010103"
            Else
                lsCuentaCorto = "24" & Mid(lsCuentaLargo, 3, 50)
            End If
        Else
            lsCuentaCorto = Me.txtCuenta.Text & Me.txtSubCuenta.Text
            If psPersCod = "1093300012530" Then 'PARA CLIENTES MI VIVIENDA
                lsCuentaLargo = "26" & Mid(lsCuentaCorto, 3, 1) & "602010103"
            Else
                lsCuentaLargo = "26" & Mid(lsCuentaCorto, 3, 50)
            End If

        End If
        
        
    
        If oCta.GetEsUltimoNivel(lsCuentaCorto) = False Or oCta.GetEsUltimoNivel(lsCuentaLargo) = False Then
            MsgBox "Cuentas de Corto / Largo no son cuentas de último nivel", vbInformation, "Aviso"
        Else
            If Len(Trim(lsCuentaCorto)) > 0 And Len(Trim(lsCuentaLargo)) > 0 Then
        
            If MsgBox("Desea Confirmar Adeudado?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                    lsMovNro = oCon.GeneraMovNro(dFechaAper, gsCodAge, gsCodUser)
                    If oAdeu.GrabaConfirmacionAdeudado(lsMovNro, gsOpeCod, txtMovDesc, _
                                                lsCuenta, lsSubCuenta, lsCtaHaber, pnImporte, psPersCod, _
                                                psTipoIF, psCtaIFCod, pnMovRef, dFechaAper, lnTasaVac, _
                                                nMontoPrestado, nMontoPrestadoReal, lsMonedaPago, _
                                                lsCuentaCorto, lsCuentaLargo) = 0 Then
                        
                        MsgBox "Registro Confirmado satisfactoriamente", vbInformation, "Aviso"
'                        If chkImpresion.value = 1 Then
                            ImprimeAsientoContable lsMovNro
'                        End If
                                            
                        MsgBox "Confirmación Efectuada Satisfactoriamente", vbInformation, "Aviso"
                    
                        Me.fgListaCG.EliminaFila Me.fgListaCG.row
                        
'                        lblCantidad.Caption = Val(lblCantidad.Caption) - 1
                        
'                        If Val(lblCantidad.Caption) = 0 Then
'                            cmdConfAper.Enabled = False
'                        Else
                            cmdConfAper.Enabled = True
'                        End If
                        
                    Else
                        MsgBox "No se pudo efectuar grabación de la confirmación del adeudo", vbInformation, "Aviso"
                    End If
                Else
                    Exit Sub
                End If
            
            Else
                MsgBox "Cuentas Contables de Corto y Largo Plazo aun no definidas para confirmación del adeudo", vbInformation, "Aviso"
            End If
        End If
    'Else
    '    MsgBox "Cuentas Contables no definidas para confirmación del adeudo", vbInformation, "Aviso"
    'End If

Exit Sub
ConfirmarError_:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"

End Sub



