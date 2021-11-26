VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmARendirExtorno 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5760
   ClientLeft      =   1275
   ClientTop       =   2205
   ClientWidth     =   9255
   Icon            =   "frmARendirExtorno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
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
      Height          =   465
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3930
      Width           =   9030
   End
   Begin VB.Frame FraRangoFechas 
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   75
      TabIndex        =   12
      Top             =   840
      Width           =   9030
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
         Height          =   330
         Left            =   7470
         TabIndex        =   5
         Top             =   180
         Width           =   1410
      End
      Begin MSMask.MaskEdBox txtDesde 
         Height          =   300
         Left            =   2565
         TabIndex        =   3
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHasta 
         Height          =   300
         Left            =   4485
         TabIndex        =   4
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
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
         Height          =   210
         Left            =   1920
         TabIndex        =   14
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label4 
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
         Height          =   210
         Left            =   3900
         TabIndex        =   13
         Top             =   225
         Width           =   540
      End
   End
   Begin VB.TextBox txtMovDescExt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4665
      Width           =   9030
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7755
      TabIndex        =   10
      Top             =   5310
      Width           =   1320
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Extornar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6390
      TabIndex        =   9
      ToolTipText     =   "Ingresar Saldo a Caja General"
      Top             =   5310
      Width           =   1320
   End
   Begin Sicmact.Usuario usu 
      Left            =   15
      Top             =   5250
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CheckBox chkSelec 
      Caption         =   "&Todos"
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
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   870
   End
   Begin Sicmact.FlexEdit fgAtenciones 
      Height          =   2430
      Left            =   75
      TabIndex        =   6
      Top             =   1440
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   4286
      Cols0           =   19
      HighLight       =   2
      AllowUserResizing=   3
      EncabezadosNombres=   $"frmARendirExtorno.frx":030A
      EncabezadosAnchos=   "350-450-1200-900-1200-0-3000-1200-0-0-0-0-0-2000-0-0-0-2000-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-C-L-C-L-R-C-R-L-L-R-L-L-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-0-0-0-0-2-0-0-0-0-0-0"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
   End
   Begin VB.Frame FraSeleccion 
      Enabled         =   0   'False
      Height          =   630
      Left            =   75
      TabIndex        =   15
      Top             =   120
      Width           =   9030
      Begin Sicmact.TxtBuscar txtBuscarAgenciaArea 
         Height          =   330
         Left            =   1410
         TabIndex        =   1
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
         lbUltimaInstancia=   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Area/Agencia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label lblAgenciaArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2550
         TabIndex        =   2
         Top             =   225
         Width           =   5940
      End
   End
   Begin VB.Label lblMensaje 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   165
      Left            =   150
      TabIndex        =   17
      Top             =   5400
      Width           =   5655
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Glosa de Operación"
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
      Left            =   90
      TabIndex        =   11
      Top             =   4425
      Width           =   1605
   End
End
Attribute VB_Name = "frmARendirExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cOpeCod As String, cOpeDesc As String
Dim lTransActiva As Boolean
Dim lRindeCajaG As Boolean, lEsChica As Boolean
Dim lRindeCajaCh As Boolean, lRindeViaticos As Boolean
Dim sObjRendir As String
Dim sDocTpoRecibo As String
Dim sCtaPendiente As String

'************************************************************************************
'************************************************************************************
Dim oContFunc As NContFunciones
Dim oAreas As DActualizaDatosArea
Dim oNArendir As NARendir
Dim oOperacion As DOperacion

Dim lnTipoArendir As ArendirTipo
Dim lbEsChica  As Boolean
Dim lsCtaArendir As String
Dim lsCtaPendiente As String
Dim lsDocTpoRecibo As String

Dim lsTpoDocVoucher  As String
Dim lSalir As Boolean
Dim lsMovNroSolicitud As String
Dim lnArendirFase As ARendirFases

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(ByVal pnTipoArendir As ArendirTipo, ByVal pnArendirFase As ARendirFases, Optional pbEsCajaChica As Boolean = False)
lnArendirFase = pnArendirFase
lnTipoArendir = pnTipoArendir
lbEsChica = pbEsCajaChica
Me.Show 1
End Sub
Private Function GetArendirAtencion() As Boolean

Dim lnFila As Long
Dim rs As ADODB.Recordset
GetArendirAtencion = False
lSalir = False
Set rs = New ADODB.Recordset
'***Comentado por ELRO el 20120410, según OYP-RFC005-2012
'If TxtBuscarArendir = "" Then
'    MsgBox "Ingrese el Area a la cual Pertenece el Arendir", vbInformation, "Aviso"
'    TxtBuscarArendir.SetFocus
'    Exit Function
'End If
'***Fin Comentado por ELRO*******************************
If chkSelec.value = 0 Then
    If txtBuscarAgenciaArea = "" Then
        MsgBox "Ingrese el Area del empleado a quien emitio el Arendir", vbInformation, "Aviso"
        txtBuscarAgenciaArea.SetFocus
        Exit Function
    End If
End If
If ValFecha(txtDesde) = False Then
    Exit Function
End If
If ValFecha(txthasta) = False Then
    Exit Function
End If
If CDate(txthasta) < CDate(txtDesde) Then
    MsgBox "Fecha Final no puede ser menor a la fecha de inicial", vbInformation, "Aviso"
    Exit Function
End If
Me.fgAtenciones.Clear
Me.fgAtenciones.FormaCabecera
Me.fgAtenciones.Rows = 2
Me.MousePointer = 11
lblMensaje = "Procesando Información. Por favor Espere... "
lblMensaje.Refresh
CabeceraFg
If lnArendirFase = ArendirExtornoAtencion Then
    '***Modificado por ELRO el 20120411, según OYP-RFC005-2012
    'Set rs = oNArendir.GetAtencionSinSustentacion(Mid(txtBuscarAgenciaArea.Text, 4, 2), Mid(txtBuscarAgenciaArea.Text, 1, 3), lnTipoArendir, lsCtaArendir, Mid(TxtBuscarArendir, 1, 3), Mid(TxtBuscarArendir, 4, 2), CDate(txtDesde), CDate(txtHasta), Mid(gsOpeCod, 3, 1))
    Set rs = oNArendir.GetAtencionSinSustentacion(Mid(txtBuscarAgenciaArea.Text, 4, 2), Mid(txtBuscarAgenciaArea.Text, 1, 3), lnTipoArendir, lsCtaArendir, "025", "", CDate(txtDesde), CDate(txthasta), Mid(gsOpeCod, 3, 1))
ElseIf lnArendirFase = 7 Then 'ArendirExtornoAprobaciónViáticos
        Set rs = oNArendir.GetAtencionSinSustentacion(IIf(Len(Trim(Mid(txtBuscarAgenciaArea.Text, 4, 2))) = 0, "01", Mid(txtBuscarAgenciaArea.Text, 4, 2)), Mid(txtBuscarAgenciaArea.Text, 1, 3), lnTipoArendir, lsCtaArendir, "", "", CDate(txtDesde), CDate(txthasta), Mid(gsOpeCod, 3, 1), chkSelec, gsOpeCod)
ElseIf lnArendirFase = 8 Then 'ArendirExtornoAprobaciónCuentas
        Set rs = oNArendir.GetAtencionSinSustentacion(IIf(Len(Trim(Mid(txtBuscarAgenciaArea.Text, 4, 2))) = 0, "01", Mid(txtBuscarAgenciaArea.Text, 4, 2)), Mid(txtBuscarAgenciaArea.Text, 1, 3), lnTipoArendir, lsCtaArendir, "", "", CDate(txtDesde), CDate(txthasta), Mid(gsOpeCod, 3, 1), chkSelec, gsOpeCod)
    '***Fin Modificado por ELRO*******************************
Else
    '***Modificado por ELRO el 20120411, según OYP-RFC005-2012
    'Set rs = oNArendir.GetRendicionesArendir(Mid(txtBuscarAgenciaArea.Text, 4, 2), Mid(txtBuscarAgenciaArea.Text, 1, 3), Mid(TxtBuscarArendir, 1, 3), Mid(TxtBuscarArendir, 4, 2), lnTipoArendir, lsCtaArendir, lsCtaPendiente, CDate(txtDesde), CDate(txthasta), Mid(gsOpeCod, 3, 1))
    Set rs = oNArendir.GetRendicionesArendir(IIf(Len(Trim(Mid(txtBuscarAgenciaArea.Text, 4, 2))) = 0, "01", Mid(txtBuscarAgenciaArea.Text, 4, 2)), Mid(txtBuscarAgenciaArea.Text, 1, 3), "", "", lnTipoArendir, lsCtaArendir, lsCtaPendiente, CDate(txtDesde), CDate(txthasta), Mid(gsOpeCod, 3, 1), chkSelec)
      
End If
If Not rs.EOF And Not rs.BOF Then
    Set fgAtenciones.Recordset = rs
    If lnArendirFase = ArendirExtornoAtencion Then
        fgAtenciones.FormatoPersNom 6
    Else
        fgAtenciones.FormatoPersNom 7
    End If
Else
    If lnTipoArendir = gArendirTipoCajaChica Then
        MsgBox "Caja Chica sin egresos pendientes de A rendir", vbInformation, "Aviso"
    Else
        MsgBox "No se encuentran A rendir Cuenta Pendientes en Rango Ingresado ", vbInformation, "Aviso"
        txtDesde.SetFocus
    End If
End If
rs.Close: Set rs = Nothing
GetArendirAtencion = True
lblMensaje = ""
lblMensaje.Refresh
Me.MousePointer = 0
End Function
Private Sub chkSelec_Click()
If chkSelec.value = 0 Then
    FraSeleccion.Enabled = True
    txtBuscarAgenciaArea.SetFocus
Else
    FraSeleccion.Enabled = False
    txtBuscarAgenciaArea.Text = ""
    lblAgenciaArea = ""
End If
End Sub
Private Sub chkSelec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FraSeleccion.Enabled Then
        txtBuscarAgenciaArea.SetFocus
    Else
        cmdProcesar.SetFocus
    End If
End If
End Sub

Private Sub cmdExtornar_Click()
Dim ldFechaAtenc As Date
Dim lsMovAtenc As String
Dim lnImporte As Currency
Dim lsDocTpo As String
Dim lsOpeDoc As String
Dim lsMovNro As String
Dim lsMovNroRend As String
Dim ldFechaRend As Date
Dim lnSaldo As Currency
Dim lsOpeRend As String
Dim lsCuentaAho As String
Dim lsNroDoc As String

If fgAtenciones.TextMatrix(1, 0) = "" Then Exit Sub
If Len(Trim(txtMovDescExt)) = 0 Then
    MsgBox "Ingrese la descripcion del extorno ", vbInformation, "Aviso"
    txtMovDescExt.SetFocus
    Exit Sub
End If

'***Agregado por ELRO el 20120524, según OYP-RFC005-2012
If fgAtenciones.TextMatrix(fgAtenciones.row, 9) = gsCodPersUser Then
    MsgBox "Usted no puede  extornar su propia operación", vbInformation, "Aviso"
    cmdExtornar.SetFocus
    Exit Sub
End If

If Left(fgAtenciones.TextMatrix(fgAtenciones.row, 5), 4) = gsCodUser Then
    MsgBox "Usted no puede  extornar su propia operación", vbInformation, "Aviso"
    cmdExtornar.SetFocus
    Exit Sub
End If

If gsOpeCod = CStr(gCGArendirViatExtRendMN) Or gsOpeCod = CStr(gCGArendirViatExtRendME) Or _
  gsOpeCod = CStr(gCGArendirCtaExtRendMN) Or gsOpeCod = CStr(gCGArendirCtaExtRendME) Then  'solo aplica para extorno de rendición
    If fgAtenciones.TextMatrix(fgAtenciones.row, 23) = CStr(gCGArendirViatRendDsctPlanMN) Or _
     fgAtenciones.TextMatrix(fgAtenciones.row, 23) = CStr(gCGArendirViatRendDsctPlanME) Or _
     fgAtenciones.TextMatrix(fgAtenciones.row, 23) = CStr(gCGArendirCtaRendDsctPlanMN) Or _
     fgAtenciones.TextMatrix(fgAtenciones.row, 23) = CStr(gCGArendirCtaRendDsctPlanME) Then
        MsgBox "No se puede extornar un A Rendir por Descuento por Planilla ", vbInformation, "Aviso"
        cmdExtornar.SetFocus
        Exit Sub
    End If
End If
'***Fin Agregado por ELRO*******************************

lsOpeDoc = gsOpeCod
If MsgBox("Desea Realizar el extorno??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
   
    Select Case lnArendirFase
        Case ArendirExtornoAtencion
            lsMovAtenc = fgAtenciones.TextMatrix(fgAtenciones.row, 10)
            lsMovNroSolicitud = fgAtenciones.TextMatrix(fgAtenciones.row, 16)
            lnImporte = CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 7))
            lnSaldo = CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 12))
            lsDocTpo = IIf(fgAtenciones.TextMatrix(fgAtenciones.row, 11) = "", "-1", fgAtenciones.TextMatrix(fgAtenciones.row, 11))
'            lsOpeDoc = oNArendir.GetOpeRendicion(Mid(gsOpeCod, 1, 5), lsDocTpo, lsCtaArendir, lsCtaPendiente)
            lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            
            oNArendir.ExtornaArendir lsMovNro, lsOpeDoc, txtMovDescExt, _
                                    lsMovAtenc, lsMovNroSolicitud, lnTipoArendir, lnImporte, lnSaldo, lnArendirFase
            
            ImprimeAsientoContable lsMovNro
        Case ArendirExtornoRendicion
            lsCuentaAho = ""
            lsMovNroRend = fgAtenciones.TextMatrix(fgAtenciones.row, 14)
            lsMovAtenc = fgAtenciones.TextMatrix(fgAtenciones.row, 15)
            lsMovNroSolicitud = fgAtenciones.TextMatrix(fgAtenciones.row, 16)
            lnImporte = CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 9))
            lnSaldo = CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 11))
            lsOpeRend = fgAtenciones.TextMatrix(fgAtenciones.row, 23)
            lsCuentaAho = fgAtenciones.TextMatrix(fgAtenciones.row, 25)
            lsNroDoc = fgAtenciones.TextMatrix(fgAtenciones.row, 3)
            lsDocTpo = IIf(nVal(fgAtenciones.TextMatrix(fgAtenciones.row, 18)) = 0, "-1", fgAtenciones.TextMatrix(fgAtenciones.row, 18))
            
            If Val(Right(lsOpeRend, 1)) = 0 Then
                lsOpeDoc = gsOpeCod
            Else
                lsOpeDoc = oNArendir.GetOpeRendicion(Mid(gsOpeCod, 1, 5), lsDocTpo, lsCtaArendir, lsCtaPendiente, IIf(Val(Right$(lsOpeRend, 1)) <= 1, False, True))
            End If
            lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            If lsCuentaAho <> "" Then
                Select Case lsDocTpo
                    Case TpoDocOrdenPago
                           oNArendir.CapExtornoCargoAhoMov lsMovNroRend, lsOpeDoc, lsCuentaAho, lsMovNro, _
                                    txtMovDescExt, Abs(lnImporte), TpoDocOrdenPago, lsNroDoc, , True, lsMovNroRend, _
                                    lsMovNroSolicitud, lnTipoArendir, lnSaldo, lnArendirFase, gdFecSis, True, gsNomAge, gsCodUser
                End Select
            Else
                oNArendir.ExtornaArendir lsMovNro, lsOpeDoc, txtMovDescExt, _
                        lsMovNroRend, lsMovNroSolicitud, lnTipoArendir, lnImporte, lnSaldo, lnArendirFase
            End If
            ImprimeAsientoContable lsMovNro
       '***Agregado por ELRO el 20120504, según OYP-REFC005-2012
       Case 7 'ArendirExtornoAprobaciónViáticos
            lsMovAtenc = fgAtenciones.TextMatrix(fgAtenciones.row, 10)
            lsMovNroSolicitud = fgAtenciones.TextMatrix(fgAtenciones.row, 16)
            lnImporte = CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 7))
            lnSaldo = CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 12))
            lsDocTpo = IIf(fgAtenciones.TextMatrix(fgAtenciones.row, 11) = "", "-1", fgAtenciones.TextMatrix(fgAtenciones.row, 11))
            lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            
            oNArendir.ExtornaArendir lsMovNro, lsOpeDoc, txtMovDescExt, _
                                    lsMovAtenc, lsMovNroSolicitud, lnTipoArendir, lnImporte, lnSaldo, lnArendirFase
            
            ImprimeAsientoContable lsMovNro
       Case 8 'ArendirExtornoAprobaciónArendir
            lsMovAtenc = fgAtenciones.TextMatrix(fgAtenciones.row, 10)
            lsMovNroSolicitud = fgAtenciones.TextMatrix(fgAtenciones.row, 16)
            lnImporte = CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 7))
            lnSaldo = CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 12))
            lsDocTpo = IIf(fgAtenciones.TextMatrix(fgAtenciones.row, 11) = "", "-1", fgAtenciones.TextMatrix(fgAtenciones.row, 11))
'            lsOpeDoc = oNArendir.GetOpeRendicion(Mid(gsOpeCod, 1, 5), lsDocTpo, lsCtaArendir, lsCtaPendiente)
            lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            
            oNArendir.ExtornaArendir lsMovNro, lsOpeDoc, txtMovDescExt, _
                                    lsMovAtenc, lsMovNroSolicitud, lnTipoArendir, lnImporte, lnSaldo, lnArendirFase
            
            ImprimeAsientoContable lsMovNro
      '***Fin Agregado por ELRO*********************************
    End Select
    fgAtenciones.EliminaFila fgAtenciones.row
    MsgBox "Extorno realizado con éxito", vbInformation, "Aviso"
                'ARLO20170208
            Set objPista = New COMManejador.Pista
            Dim lsArendir As String
            If (gsOpeCod = 401180) Then
            lsArendir = fgAtenciones.TextMatrix(fgAtenciones.row, 5)
            Else
            lsArendir = fgAtenciones.TextMatrix(fgAtenciones.row, 4)
            End If
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " N° A Rendir : " & lsArendir & " Motivo : " & txtMovDescExt.Text
            Set objPista = Nothing
            '*******
    txtMovDescExt = ""
    txtMovDesc = ""
End If
End Sub
Private Sub cmdProcesar_Click()
Dim lbDatos As Boolean
Dim oCont As New NContFunciones
If Not oCont.PermiteModificarAsiento(Format(Me.txtDesde, gsFormatoMovFecha), False) Then
   MsgBox "No se puede Extornar operación de Mes Contable Cerrado!", vbInformation, "¡Aviso!"
   Exit Sub
End If
Set oCont = Nothing

If GetArendirAtencion Then
    fgAtenciones.SetFocus
Else
    txtBuscarAgenciaArea.Text = ""
    lblAgenciaArea = ""
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub fgAtenciones_Click()
If fgAtenciones.TextMatrix(1, 0) <> "" Then
    If lnArendirFase = ArendirExtornoRendicion Then
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 17)
    Else
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 8)
    End If
End If
End Sub
Private Sub fgAtenciones_GotFocus()
If fgAtenciones.TextMatrix(1, 0) <> "" Then
    If lnArendirFase = ArendirExtornoRendicion Then
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 17)
    Else
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 8)
    End If
End If

End Sub
Private Sub fgAtenciones_OnRowChange(pnRow As Long, pnCol As Long)
If fgAtenciones.TextMatrix(1, 0) <> "" Then
    If lnArendirFase = ArendirExtornoRendicion Then
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 17)
    Else
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 8)
    End If
End If
End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
End Sub
Private Sub Form_Load()
Dim lvItem As ListItem
Dim rsPer As New ADODB.Recordset
Dim sOpeCod As String
 
Set oContFunc = New NContFunciones
Set oNArendir = New NARendir
Set oAreas = New DActualizaDatosArea
Set oOperacion = New DOperacion
lSalir = False
Me.Caption = gsOpeDesc

CentraForm Me
'***Comentado por ELRO el 20120410. según OYP-RFC005-2012
'TxtBuscarArendir.psRaiz = "A Rendir de..."

'If lnTipoArendir = gArendirTipoAgencias Then
'    Set rsPer = oOperacion.CargaOpeObj(gCGArendirCtaSolMNAge, 1)
'Else
'    Set rsPer = oOperacion.CargaOpeObj(gCGArendirCtaSolMN, 1)
'End If
'If Not rsPer.EOF Then
'    TxtBuscarArendir.rs = oAreas.GetAgenciasAreas(rsPer!cOpeObjFiltro, 1)
'End If
'RSClose rsPer
'***Fin Comentado por ELRO*******************************

If Mid(gsOpeCod, 3, 1) = gMonedaNacional Then
   gsSimbolo = gcMN
Else
   gsSimbolo = gcME
End If
txtDesde = gdFecSis
txthasta = gdFecSis
lsTpoDocVoucher = oOperacion.EmiteDocOpe(gsOpeCod, OpeDocEstOpcionalDebeExistir, OpeDocMetAutogenerado)
lsCtaArendir = oOperacion.EmiteOpeCta(gsOpeCod, "H", "0")
If lsCtaArendir = "" Then
   MsgBox "Faltan asignar Cuenta Contable de Arendir a Operación." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Sistemas", vbInformation, "Aviso"
   lSalir = True
   Exit Sub
End If
lsCtaPendiente = oOperacion.EmiteOpeCta(gsOpeCod, "H", "1")
If lsCtaPendiente = "" Then
   MsgBox "Falta asignar Cuenta de Pendiente a Operación." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Sistemas", vbInformation, "Aviso"
   lSalir = True
   Exit Sub
End If
If lnTipoArendir = gArendirTipoCajaChica Then
    txtBuscarAgenciaArea.psRaiz = "CAJAS CHICAS"
    txtBuscarAgenciaArea.rs = oNArendir.EmiteCajasChicas
Else
    txtBuscarAgenciaArea.rs = oAreas.GetAgenciasAreas
End If
If lnArendirFase = ArendirExtornoAtencion Then
    FraRangoFechas.Caption = "Fecha de Atención"
'***Agregado por ELRO el 20120411, según OYP-RFC005-2012
ElseIf lnArendirFase = 7 Or lnArendirFase = 8 Then  'ArendirExtornoAprobaciónViaticos y ArendirExtornoAprobaciónARendir
    FraRangoFechas.Caption = "Fecha de Aprobación"

'***Fin Agregado por ELRO*******************************
Else
    FraRangoFechas.Caption = "Fecha de Rendición"
End If
Select Case lnTipoArendir
    Case gArendirTipoCajaChica
         Me.Height = 5550
         cmdSalir.Top = 5030 - cmdSalir.Height
         lsDocTpoRecibo = oOperacion.EmiteDocOpe(gsOpeCod, OpeDocEstObligatorioDebeExistir, OpeDocMetDigitado)
         If lsDocTpoRecibo = "" Then
            MsgBox "No se asignó Tipo de Documento Recibo de A rendir a Operación", vbInformation, "Error"
            lSalir = True
            Exit Sub
         End If
         fgAtenciones.ColWidth(2) = fgAtenciones.ColWidth(2) - 300
         fgAtenciones.ColWidth(3) = fgAtenciones.ColWidth(3) + 300
End Select
CabeceraFg
End Sub
Private Sub CabeceraFg()
If lnArendirFase = ArendirExtornoAtencion Then
    fgAtenciones.ListaControles = "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
    fgAtenciones.FormatosEdit = "0-0-0-0-0-0-0-2-0-2-0-0-2-0-0-0-0-0-0"
    fgAtenciones.EncabezadosNombres = " -Tipo-Número-Fecha-N° ARendir-Fecha-Persona-Importe-cMovDesc-cCodPers-cMovNro-nDocTpo-Saldo-Area Funcional-cAgecod-cDocDes-MovNroSolicitud-Agencia-cAgeCod"
    fgAtenciones.EncabezadosAnchos = "350-450-1200-0-1200-900-3000-1200-0-0-0-0-0-2000-0-0-0-2000-0-0-0-0-0"
    fgAtenciones.EncabezadosAlineacion = "C-C-L-C-L-C-L-R-C-R-L-L-R-L-L-C-C-C-C"
'***Agregado por ELRO el 20120504, según OYP-RFC005-2012
ElseIf lnArendirFase = 7 Then 'ArendirExtornoAprobaciónViáticos
    fgAtenciones.ListaControles = "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
    fgAtenciones.FormatosEdit = "0-0-0-0-0-0-0-2-0-0-0-0-2-0-0-0-0-0-0"
    fgAtenciones.EncabezadosNombres = " -Tipo-Número-Fecha-N° ARendir-Fecha-Persona-Importe-cMovDesc-cCodPers-cMovNro-nDocTpo-Saldo-Area Funcional-cAgecod-cDocDes-MovNroSolicitud-Agencia-cAgeCod"
    fgAtenciones.EncabezadosAnchos = "350-450-1200-0-1200-900-3000-1200-0-0-0-0-0-2000-0-0-0-2000-0-0-0-0-0"
    fgAtenciones.EncabezadosAlineacion = "C-C-L-C-L-C-L-R-C-R-L-L-R-L-L-C-C-C-C"
ElseIf lnArendirFase = 8 Then 'ArendirExtornoAprobaciónARendir
    fgAtenciones.ListaControles = "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
    fgAtenciones.FormatosEdit = "0-0-0-0-0-0-0-2-0-0-0-0-2-0-0-0-0-0-0"
    fgAtenciones.EncabezadosNombres = " -Tipo-Número-Fecha-N° ARendir-Fecha-Persona-Importe-cMovDesc-cCodPers-cMovNro-nDocTpo-Saldo-Area Funcional-cAgecod-cDocDes-MovNroSolicitud-Agencia-cAgeCod"
    fgAtenciones.EncabezadosAnchos = "350-450-1200-0-1200-900-3000-1200-0-0-0-0-0-2000-0-0-0-2000-0-0-0-0-0"
    fgAtenciones.EncabezadosAlineacion = "C-C-L-C-L-C-L-R-C-R-L-L-R-L-L-C-C-C-C"
'***Fin Agregado por ELRO*******************************
Else
    fgAtenciones.ListaControles = "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
    fgAtenciones.FormatosEdit = "0-0-0-0-0-0-0-2-2-2-0-0-0-0-0-0-0-0-0-0-0-0-0"
    fgAtenciones.EncabezadosNombres = " -Rendicion -Tipo-Número-Fecha-ARendir-Fecha-Persona-Monto-Monto Real-MontoAtenc-Saldo " _
                                     & "Area Funcional- Area -cMovRendicion-cMovNroAtenc-cMovNroSol-GlosaRend-nDocTpoRend-nDocTpoSol-" _
                                     & "cAreaCod-cAgeCod-cAgeDescripcion-cOpeCod"

    fgAtenciones.EncabezadosAnchos = "350-1900-500-800-800-1100-800-2500-1200-0-0-0-2000-0-0-0-0-0-0-0-0-0-0-0-0"
    fgAtenciones.EncabezadosAlineacion = "C-L-C-C-C-L-C-L-R-R-R-L-L-L-L-L-L-L-L"
End If
fgAtenciones.Clear
fgAtenciones.Rows = 2
fgAtenciones.FormaCabecera

End Sub
Private Sub Form_Unload(Cancel As Integer)
Set oContFunc = Nothing
Set oNArendir = Nothing
Set oAreas = Nothing
End Sub
Public Property Get sPendiente() As String
sPendiente = sCtaPendiente
End Property
Public Property Let sPendiente(ByVal vNewValue As String)
sCtaPendiente = sPendiente
End Property


Private Sub txtBuscarAgenciaArea_EmiteDatos()
lblAgenciaArea = oAreas.GetNombreAreas(Mid(txtBuscarAgenciaArea, 1, 3)) + " " + oAreas.GetNombreAgencia(Mid(txtBuscarAgenciaArea, 4, 2))
'lblAgeDesc = oAreas.GetNombreAgencia(Mid(txtBuscarAgenciaArea, 4, 2))
If txtDesde.Visible Then
    txtDesde.SetFocus
End If
End Sub
'***Comentado por ELRO el 20120410, según OYP-RFC005-2012
'Private Sub TxtBuscarArendir_EmiteDatos()
'lblDescArendir = TxtBuscarArendir.psDescripcion
'If chkSelec.Visible Then
'    chkSelec.SetFocus
'End If
'End Sub
'***Fin Comentado por ELRO*******************************

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txthasta.SetFocus
End Sub
Private Sub txthasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdProcesar.SetFocus
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtMovDescExt.SetFocus
End Sub
Private Sub txtMovDescExt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdExtornar.SetFocus
End If
End Sub
