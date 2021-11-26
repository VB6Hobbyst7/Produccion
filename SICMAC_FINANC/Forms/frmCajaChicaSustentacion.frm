VERSION 5.00
Begin VB.Form frmCajaChicaSustentacion 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   Icon            =   "frmCajaChicaSustentacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRendicion 
      Caption         =   "&Rendición"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   23
      Top             =   4560
      Width           =   1245
   End
   Begin VB.CheckBox chkTodo 
      Caption         =   "&Todo"
      Height          =   255
      Left            =   8280
      TabIndex        =   22
      Top             =   840
      Value           =   1  'Checked
      Width           =   735
   End
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
      Height          =   360
      Left            =   8280
      TabIndex        =   21
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Frame fraColaboradores 
      Caption         =   "Colaborador"
      Enabled         =   0   'False
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
      Height          =   1785
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   8100
      Begin Sicmact.TxtBuscar txtBuscaPers 
         Height          =   330
         Left            =   1020
         TabIndex        =   11
         Top             =   240
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
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
         TabIndex        =   20
         Top             =   720
         Width           =   750
      End
      Begin VB.Label lblpersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1020
         TabIndex        =   19
         Top             =   600
         Width           =   6120
      End
      Begin VB.Label lblAgeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1620
         TabIndex        =   18
         Top             =   1320
         Width           =   5535
      End
      Begin VB.Label lblAreaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1620
         TabIndex        =   17
         Top             =   960
         Width           =   5520
      End
      Begin VB.Label lblAgecod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1020
         TabIndex        =   16
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label lblAreaCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1020
         TabIndex        =   15
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
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
         TabIndex        =   14
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Area :"
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
         TabIndex        =   13
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Persona :"
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
         TabIndex        =   12
         Top             =   270
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8325
      TabIndex        =   9
      Top             =   4560
      Width           =   1245
   End
   Begin VB.CommandButton cmdSustentacion 
      Caption         =   "Sus&tentación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7080
      TabIndex        =   8
      Top             =   4560
      Width           =   1245
   End
   Begin VB.Frame Frame1 
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
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   9540
      Begin Sicmact.FlexEdit fgListaCH 
         Height          =   1515
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2672
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Nro Doc.-Fecha-Solicitante-Importe-cCodPers-Glosa-Area-cMovNroArendir-cCodArea"
         EncabezadosAnchos=   "450-1-1200-4000-1200-0-0-1200-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-L-R-L-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-2-0-0-0-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
         RowHeightMin    =   150
      End
   End
   Begin VB.Frame fraCajaChica 
      Caption         =   "Caja Chica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8100
      Begin Sicmact.TxtBuscar txtBuscarAreaCH 
         Height          =   345
         Left            =   1200
         TabIndex        =   1
         Top             =   225
         Width           =   1095
         _ExtentX        =   1535
         _ExtentY        =   609
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
      Begin VB.Label lblNroProcCH 
         Alignment       =   2  'Center
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
         Height          =   300
         Left            =   7260
         TabIndex        =   5
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "N° :"
         Height          =   195
         Left            =   6915
         TabIndex        =   4
         Top             =   300
         Width           =   270
      End
      Begin VB.Label lblCajaChicaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2310
         TabIndex        =   3
         Top             =   225
         Width           =   4440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Area/Agencia : "
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   300
         Width           =   1125
      End
   End
   Begin Sicmact.Usuario Usu 
      Left            =   120
      Top             =   4560
      _ExtentX        =   820
      _ExtentY        =   820
   End
End
Attribute VB_Name = "frmCajaChicaSustentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
'***Nombre:         frmCajaChicaSustentacion
'***Descripción:    Formulario para Sustentar la Entrega de Efectivo de Caja Chica
'***Creación:       ELRO el 20120614, según OYP-RFC047-2012
'*********************************************************************************
Option Explicit
Dim oArendir As NARendir
Dim oCH As NCajaChica
Dim lnArendirFase As ARendirFases
Dim lsCtaArendir As String
Dim lsCtaPendiente As String
Dim lsCtaFondofijo As String
Dim lbSalir As Boolean
Public Sub Inicio(Optional ByVal pnArendirFase As ARendirFases = ArendirRechazo)
lnArendirFase = pnArendirFase
Me.Show 1
End Sub

Private Sub limpiarCampos()
    txtBuscaPers = ""
    lblPersNombre = ""
    lblAreaCod = ""
    lblAreaDesc = ""
    lblAgecod = ""
    lblAgeDesc = ""
End Sub


Private Sub chkTodo_Click()
 If chkTodo Then
    fraColaboradores.Enabled = False
    limpiarCampos
 Else
    fraColaboradores.Enabled = True
    Call LimpiaFlex(fgListaCH)
  End If
End Sub

Private Sub asignarValores()
    txtBuscaPers.Text = Usu.PersCod
    lblAreaCod = Usu.AreaCod
    lblAgecod = Usu.CodAgeAct
    lblAgeDesc = Usu.DescAgeAct
    lblAreaDesc = Usu.AreaNom
    lblPersNombre = PstaNombre(Usu.UserNom)
End Sub

Private Sub cmdProcesar_Click()
Dim rs As ADODB.Recordset
Dim lnTipoRend As RendicionTipo
Set rs = New ADODB.Recordset

Me.MousePointer = 11
Select Case lnArendirFase
    Case ArendirRechazo, ArendirAtencion
        Set rs = oCH.GetSolicitudesArendirCH(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH))
    Case ArendirSustentacion, ArendirRendicion
        Set rs = oCH.GetCHSustSinRendicion(Mid(txtBuscarAreaCH, 1, 3), _
                                           Mid(txtBuscarAreaCH, 4, 2), _
                                           Val(lblNroProcCH), _
                                           lsCtaArendir, _
                                           IIf(lnArendirFase = ArendirRendicion, 0, chkTodo), _
                                           IIf(lnArendirFase = ArendirRendicion, "", txtBuscaPers))
    Case ArendirExtornoAtencion
        Set rs = oCH.GetAtencionesArendir(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaArendir)
    Case ArendirExtornoRendicion
        Select Case gsOpeCod
            Case gCHArendirCtaRendExacMN, gCHArendirCtaRendExacME
                lnTipoRend = Exacta
            Case gCHArendirCtaRendExtIngMN, gCHArendirCtaRendExtIngME
                lnTipoRend = ConIngreso
            Case gCHArendirCtaRendExtEgrMN, gCHArendirCtaRendExtEgrME
                lnTipoRend = ConEgreso
        End Select
        Set rs = oCH.GetCHRendiciones(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), Val(lblNroProcCH), lsCtaArendir, lsCtaPendiente, lnTipoRend)
End Select
If Not rs.EOF And Not rs.BOF Then
    Set fgListaCH.Recordset = rs
    fgListaCH.FormatoPersNom (3)
    If fgListaCH.Enabled Then
        Me.fgListaCH.SetFocus
    End If
Else
    MsgBox "No se encuentran solicitudes pendientes de Arendir", vbInformation, "Aviso"
    If txtBuscarAreaCH.Enabled And fraCajaChica.Enabled Then
        Me.txtBuscarAreaCH.SetFocus
    Else
        cmdSalir.SetFocus
    End If
End If
rs.Close
Set rs = Nothing
Me.MousePointer = 0
End Sub

Private Sub cmdRendicion_Click()
Dim oDOperacion As DOperacion
Set oDOperacion = New DOperacion
Dim oContFunc As NContFunciones
Set oContFunc = New NContFunciones
Dim lnSaldo As Currency
Dim lsMovNroAtenc As String, lsMovNroSol As String
Dim lsMovNro As String, lsOpeCod As String
Dim lnSaldoCH As Currency
Dim lsPersCod As String
Dim lsTexto As String
Dim vCtaFondoFijo As String
Dim lsSubCta As String
Dim lsCtaArendir2 As String
Dim lsCtaFondofijo2 As String
Dim lsGlosa As String '***Agregado por ELRO el 20130222, según SATI INC1301300007

On Error GoTo cmdRendicionErr

If fgListaCH.TextMatrix(1, 0) = "" Then Exit Sub

lnSaldo = CCur(fgListaCH.TextMatrix(fgListaCH.Row, 4))
lsMovNroAtenc = fgListaCH.TextMatrix(fgListaCH.Row, 10)
lsMovNroSol = fgListaCH.TextMatrix(fgListaCH.Row, 9)
lsPersCod = fgListaCH.TextMatrix(fgListaCH.Row, 14)

lsCtaArendir2 = oDOperacion.EmiteOpeCta("401350", "H")
lsCtaFondofijo2 = oDOperacion.EmiteOpeCta("401350", "D")

lsTexto = ""
If MsgBox(" ¿ Seguro de Realizar la Rendicion del A Rendir N° :" & fgListaCH.TextMatrix(fgListaCH.Row, 1) & vbCrLf & vbCrLf & "Solicitado por :" & fgListaCH.TextMatrix(fgListaCH.Row, 6), vbQuestion + vbYesNo, "Confirmación") = vbYes Then
    lsMovNro = oContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    lnSaldoCH = oCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), SaldoActual)
    Select Case lnSaldo
        Case 0  'RENDICION EXACTA
            lsOpeCod = IIf(Mid(gsOpeCod, 3, 1) = "1", gCHArendirCtaRendExacMN, gCHArendirCtaRendExacME)
            lnSaldoCH = 0
            oArendir.GrabaRendicionExacta gArendirTipoCajaChica, _
                                          gsFormatoFecha, _
                                          lsMovNro, _
                                          lsOpeCod, _
                                          "RENDICION EXACTA", _
                                          lsMovNroAtenc, _
                                          lsMovNroSol, _
                                          Mid(txtBuscarAreaCH, 1, 3), _
                                          Mid(txtBuscarAreaCH, 4, 2), _
                                          Val(lblNroProcCH), _
                                          lnSaldoCH

        Case Is > 0 'INGRESO CON EFECTIVO
            lsOpeCod = IIf(Mid(gsOpeCod, 3, 1) = "1", gCHArendirCtaRendIngMN, gCHArendirCtaRendIngME)
            '***Comentado por ELRO el 20130222, según SATI INC1301300007
            'lsSubCta = oContFunc.GetFiltroObjetos(ObjCMACAgenciaArea, lsCtaFondofijo2, Trim(txtBuscarAreaCH), False)
            'If lsSubCta <> "" Then lsCtaFondofijo2 = lsCtaFondofijo2 + IIf(CCur(lsSubCta) > 90, "01", "02") & lsSubCta
            '***Comentado por ELRO el 20130222**************************
            lsGlosa = UCase("INGRESO POR RENDICIÓN DE LA CAJA CHICA " & lblCajaChicaDesc) '***Agregado por ELRO el 20130222, según SATI INC1301300007
            vCtaFondoFijo = lsCtaFondofijo2
            
            oArendir.GrabaRendicionGiroDocumento gArendirTipoCajaChica, _
                                                 lsMovNro, lsMovNroSol, _
                                                 lsMovNroAtenc, _
                                                 lsOpeCod, lsGlosa, _
                                                 vCtaFondoFijo, lsCtaArendir2, _
                                                 lsPersCod, Abs(lnSaldo), _
                                                 TpoDocRecEgreso, _
                                                 "", gdFecSis, "", "", "", False, _
                                                 Mid(txtBuscarAreaCH, 1, 3), _
                                                 Mid(txtBuscarAreaCH, 4, 2), _
                                                 Val(lblNroProcCH)
            '***Parametro lsGlosa agregado por ELRO el 20130222, según SATI INC1301300007
    End Select
    fgListaCH.EliminaFila fgListaCH.Row
    fgListaCH.SetFocus
End If
Set oContFunc = Nothing
Set oDOperacion = Nothing
Exit Sub
cmdRendicionErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub cmdSustentacion_Click()
Dim lsNroArendir As String
Dim lsNroDoc As String
Dim lsFechaDoc As String
Dim lsPersCod As String
Dim lsPersNomb As String
Dim lsAreaCod As String
Dim lsAreaDesc As String
Dim lsDescDoc As String
Dim lnImporte As Currency
Dim lnSaldo As Currency
Dim lsMovNroAtenc As String
Dim lsMovNroSolicitud As String
Dim lsAgeCod As String
Dim lsAgeDesc As String
Dim lsAreaCh As String
Dim lsAgeCh As String
Dim lnNroProc As Integer
Dim lsGlosa As String '***Agregado por ELRO el 20130221,según SATI INC1301300007

If fgListaCH.TextMatrix(1, 0) = "" Then Exit Sub

lsNroArendir = fgListaCH.TextMatrix(fgListaCH.Row, 1)
lsNroDoc = fgListaCH.TextMatrix(fgListaCH.Row, 1)
lsFechaDoc = fgListaCH.TextMatrix(fgListaCH.Row, 2)
lsPersCod = fgListaCH.TextMatrix(fgListaCH.Row, 14)
lsPersNomb = fgListaCH.TextMatrix(fgListaCH.Row, 6)
lsAreaCod = fgListaCH.TextMatrix(fgListaCH.Row, 11)
lsAreaDesc = fgListaCH.TextMatrix(fgListaCH.Row, 8)

lsDescDoc = fgListaCH.TextMatrix(fgListaCH.Row, 1)
lnImporte = CCur(fgListaCH.TextMatrix(fgListaCH.Row, 3))
lnSaldo = CCur(fgListaCH.TextMatrix(fgListaCH.Row, 4))
lsMovNroAtenc = fgListaCH.TextMatrix(fgListaCH.Row, 10)
lsMovNroSolicitud = fgListaCH.TextMatrix(fgListaCH.Row, 9)
lsAgeDesc = fgListaCH.TextMatrix(fgListaCH.Row, 12)
lsAgeCod = fgListaCH.TextMatrix(fgListaCH.Row, 13)
lsAreaCh = Mid(txtBuscarAreaCH, 1, 3)
lsAgeCh = Mid(txtBuscarAreaCH, 4, 2)
lnNroProc = Val(lblNroProcCH)
lsGlosa = fgListaCH.TextMatrix(fgListaCH.Row, 7)
'***Modificado por ELRO el 20120928, según OYP-RFC111-2012
'frmOpeRegDocs.Inicio lnArendirFase, gArendirTipoCajaChica, False, lsNroArendir, lsNroDoc, lsFechaDoc, lsPersCod, _
'                     lsPersNomb, lsAreaCod, lsAreaDesc, lsAgeCod, lsAgeDesc, lsDescDoc, lsMovNroAtenc, lnImporte, lsCtaArendir, _
'                     lsCtaPendiente, lnSaldo, lsMovNroSolicitud, lsAreaCh, lsAgeCh, lnNroProc
frmOpeRegDocs.Inicio lnArendirFase, gArendirTipoCajaChica, True, lsNroArendir, lsNroDoc, lsFechaDoc, lsPersCod, _
                     lsPersNomb, lsAreaCod, lsAreaDesc, lsAgeCod, lsAgeDesc, lsDescDoc, lsMovNroAtenc, lnImporte, lsCtaArendir, _
                     lsCtaPendiente, lnSaldo, lsMovNroSolicitud, lsAreaCh, lsAgeCh, lnNroProc, , lsGlosa
'***Fin Modificado por ELRO el 20120928******************
fgListaCH.TextMatrix(fgListaCH.Row, 4) = Format(frmOpeRegDocs.lnSaldo, gsFormatoNumeroView)
fgListaCH.SetFocus

End Sub



Private Sub Form_Activate()
    If lbSalir Then
        Unload Me
    End If
    txtBuscarAreaCH_EmiteDatos
    cmdProcesar_Click
End Sub

Private Sub Form_Load()
Set oArendir = New NARendir
Set oCH = New NCajaChica
Dim oOpe As DOperacion

lbSalir = False
Me.Caption = gsOpeDesc
Set oOpe = New DOperacion


cmdSustentacion.Visible = False
cmdRendicion.Visible = False

Select Case lnArendirFase

    Case ArendirSustentacion
        Me.cmdSustentacion.Visible = True
        fgListaCH.EncabezadosNombres = "N°-Nro Doc.-Fecha-Importe-Saldo-Usuario-Solicitante-Glosa-Area-cMovSol-cMovNroAtenc-cAreaCod-Agencia-cAgeCod-cPersCod-nProcNro"
        fgListaCH.EncabezadosAnchos = "450-0-1100-800-800-800-3000-4500-0-0-0-0-0-0-0-0"
        fgListaCH.EncabezadosAlineacion = "C-L-R-R-C-C-L-L-L-L-L-L-L-L-L-L"
        fgListaCH.FormatosEdit = "0-0-0-2-2-2-0-0-0-0-0-0-0-0-0-0"
        lsCtaArendir = oOpe.EmiteOpeCta(gsOpeCod, "H")
        lsCtaPendiente = oOpe.EmiteOpeCta(gsOpeCod, "H", 1)
        cmdRendicion.Visible = True
        If Trim(lsCtaArendir) = "" Or Trim(lsCtaPendiente) = "" Then
            MsgBox "Cuentas Contables de Operación no se han definido ", vbInformation, "Aviso"
            lbSalir = True
            Exit Sub
        End If
    Case ArendirRendicion
        fgListaCH.EncabezadosNombres = "N°-Nro Doc.-Fecha Doc-Solicitante-Importe - Saldo -Glosa-Area-cMovSol -cMovNroAtenc-cAreaCod- Agencia- cAgeCod - cPersCod"
        fgListaCH.EncabezadosAnchos = "450-1300-1100-3000-1200-1200-0-2000-0-0-0-2000-0-0"
        fgListaCH.EncabezadosAlineacion = "C-L-C-L-R-R-L-L-L-C-C-L-L"
        fgListaCH.FormatosEdit = "0-0-0-0-2-2-0-0-0-0-0"
        cmdRendicion.Visible = True
        cmdSustentacion.Visible = True
        lsCtaArendir = oOpe.EmiteOpeCta(gsOpeCod, "H")
        lsCtaPendiente = oOpe.EmiteOpeCta(gsOpeCod, "D", 1)
        lsCtaFondofijo = oOpe.EmiteOpeCta(gsOpeCod, "D")
        If Trim(lsCtaArendir) = "" Or Trim(lsCtaPendiente) = "" Or Trim(lsCtaFondofijo) = "" Then
            MsgBox "Cuentas Contables de Operación no se han definido ", vbInformation, "Aviso"
            lbSalir = True
            Exit Sub
        End If


End Select
CentraForm Me
txtBuscarAreaCH.psRaiz = "CAJAS CHICAS"
txtBuscarAreaCH.rs = oArendir.EmiteCajasChicas
'***Agregado por ELRO el 20120623, según OYP-RFC047-2012
fraCajaChica.Enabled = False
verificarEncargadoCH
If Len(Trim(txtBuscarAreaCH)) = 0 Then
    Exit Sub
End If
'***Fin Agregado por ELRO*******************************
Set oOpe = Nothing
End Sub


Private Sub txtBuscaPers_EmiteDatos()
    If txtBuscaPers.Text = "" Then Exit Sub
    Usu.DatosPers txtBuscaPers.Text
    If Usu.PersCod = "" Then
        MsgBox "Persona no Válida o no se Encuentra Registrada como Trabajador en la Institución", vbInformation, "Aviso"
    End If
    asignarValores
End Sub

Private Sub txtBuscarAreaCH_EmiteDatos()
Dim oCajaCH As NCajaChica
Set oCajaCH = New NCajaChica
lblCajaChicaDesc = txtBuscarAreaCH.psDescripcion
lblNroProcCH = oCajaCH.GetDatosCajaChica(Mid(txtBuscarAreaCH, 1, 3), Mid(txtBuscarAreaCH, 4, 2), NroCajaChica)
fgListaCH.Clear
fgListaCH.FormaCabecera
fgListaCH.Rows = 2
If lblCajaChicaDesc <> "" Then
    If cmdProcesar.Visible Then
       cmdProcesar.SetFocus
    End If
End If
Set oCajaCH = Nothing
End Sub

Private Sub txtBuscarAreaCH_Validate(Cancel As Boolean)
If txtBuscarAreaCH = "" Then
    Cancel = True
End If
End Sub

Private Sub verificarEncargadoCH()
    Dim oNCajaChica As NCajaChica
    Set oNCajaChica = New NCajaChica
    Dim rsEncargado As ADODB.Recordset
    Set rsEncargado = New ADODB.Recordset
    
    Set rsEncargado = oNCajaChica.verificarEncargadoCH(gsCodPersUser)
    
    If Not rsEncargado.BOF And Not rsEncargado.EOF Then
        txtBuscarAreaCH = rsEncargado!cAreaCod & rsEncargado!cAgecod
    Else
        MsgBox "No carga el código de la Caja Chica por los siguientes motivos:" & Chr(10) & "1. No esta encargado de la Caja Chica." & Chr(10) & "2. Aún no esta Autorizado el nuevo proceso de la Caja Chica." & Chr(10) & "3. Aún no cobra el efectivo habilitado por la Caja Chica.", vbInformation, "Aviso"
    End If
    Set rsEncargado = Nothing
    Set oNCajaChica = Nothing
End Sub

