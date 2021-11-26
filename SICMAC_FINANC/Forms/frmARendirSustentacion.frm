VERSION 5.00
Begin VB.Form frmARendirSustentacion 
   Caption         =   "Sustentar Gasto"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   Icon            =   "frmARendirSustentacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraRecViaticos 
      Caption         =   "Colaborador"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   7335
      Begin Sicmact.TxtBuscar txtBuscaPers 
         Height          =   300
         Left            =   1020
         TabIndex        =   7
         Top             =   240
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
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
      Begin VB.Label lblCodCat 
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
         Height          =   300
         Left            =   1020
         TabIndex        =   23
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblCategoria 
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
         Height          =   300
         Left            =   1620
         TabIndex        =   22
         Top             =   1560
         Width           =   1605
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
         TabIndex        =   21
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cargo :"
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
         Left            =   3240
         TabIndex        =   20
         Top             =   1590
         Width           =   585
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
         TabIndex        =   19
         Top             =   960
         Width           =   480
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
         TabIndex        =   18
         Top             =   1290
         Width           =   750
      End
      Begin VB.Label lblDesCargo 
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
         Height          =   300
         Left            =   3870
         TabIndex        =   17
         Top             =   1560
         Width           =   3285
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
         Height          =   300
         Left            =   1020
         TabIndex        =   16
         Top             =   900
         Width           =   585
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
         Height          =   300
         Left            =   1020
         TabIndex        =   15
         Top             =   1230
         Width           =   585
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
         Height          =   300
         Left            =   1620
         TabIndex        =   14
         Top             =   900
         Width           =   5520
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
         Height          =   300
         Left            =   1620
         TabIndex        =   13
         Top             =   1230
         Width           =   5535
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Categoria :"
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
         Top             =   1620
         Width           =   885
      End
      Begin VB.Label lblNrodoc 
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
         Height          =   300
         Left            =   5400
         TabIndex        =   11
         Top             =   240
         Width           =   1695
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
         Height          =   300
         Left            =   1020
         TabIndex        =   10
         Top             =   570
         Width           =   6120
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "L.E./DNI. :"
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
         Left            =   4680
         TabIndex        =   9
         Top             =   270
         Width           =   720
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
         TabIndex        =   8
         Top             =   615
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdSustentar 
      Caption         =   "S&ustentar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      ToolTipText     =   "Regularizar con Documentos sustentatorios"
      Top             =   3840
      Width           =   1260
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
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   3840
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      Caption         =   "Glosa"
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
      Height          =   915
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   4680
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
         Height          =   600
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   195
         Width           =   4410
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   7335
      Begin Sicmact.FlexEdit FEAtenciones 
         Height          =   1065
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   1879
         Cols0           =   10
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Tipo-N° ARendir-Fecha Solicitud-Importe-Saldo-Glosa-nMovNroVia-cDocDesc-nMovNroAte"
         EncabezadosAnchos=   "350-450-1100-1200-1000-1000-0-0-0-1"
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
            Name            =   "Tahoma"
            Size            =   6.75
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
         EncabezadosAlineacion=   "C-C-L-C-R-R-C-C-R-C"
         FormatosEdit    =   "0-0-0-0-4-2-0-0-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
   End
   Begin Sicmact.Usuario user 
      Left            =   4800
      Top             =   4080
      _ExtentX        =   820
      _ExtentY        =   820
   End
End
Attribute VB_Name = "FrmARendirSustentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************************
'***Nombre      : FrmARendirSustentacion
'***Descripción : Formulario para Sustentar los Váticos y Otros Gastos A Rendir.
'***Creación    : ELRO el 20120423
'*******************************************************************************


Private lnArendirFase As Integer
Private lnTipoArendir As Integer
Private fnSalir As Boolean
Private fsCtaArendir, fsCtaPendiente  As String
Dim lnMovNro As Long 'PASI20140507 TI-ERS060-2014

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub iniciarSustentacion(ByVal pnTipoArendir As ArendirTipo, ByVal pnArendirFase As ARendirFases)
    lnArendirFase = pnArendirFase
    lnTipoArendir = pnTipoArendir
    Me.Caption = gsOpeDesc
    fnSalir = False
    Show 1
End Sub

Private Sub actualizarAtencion()
If FEAtenciones.TextMatrix(1, 1) <> "" Then
    txtMovDesc = FEAtenciones.TextMatrix(FEAtenciones.row, 6)
Else
    txtMovDesc = ""
End If
End Sub

Private Sub limpiarCampos()
txtBuscaPers = ""
lblNrodoc = ""
lblPersNombre = ""
lblAreaCod = ""
lblAreaDesc = ""
lblAgeCod = ""
lblAgeDesc = ""
lblCodCat = ""
lblCategoria = ""
lblDesCargo = ""
End Sub

Private Function devolverDatosColaborador(psUser As String, Optional psPersCod As String = "") As Boolean

devolverDatosColaborador = False

If psUser <> "" Then
    user.Inicio gsCodUser
    
Else
    user.DatosPers psPersCod
End If

If Trim(user.PersCod) <> "" Then
    txtBuscaPers.Text = user.PersCod
     txtBuscaPers.psCodigoPersona = user.PersCod
    lblPersNombre = PstaNombre(user.UserNom)
    lblAgeCod = user.CodAgeAct
    lblAgeDesc = user.DescAgeAct
    lblAreaCod = user.AreaCod
    lblAreaDesc = user.AreaNom
    lblDesCargo = user.PersCargo
    lblNrodoc = user.NroDNIUser
    If user.PersCategCod <> "" Then
        lblCategoria = user.PersCategDesc
        lblCodCat = user.PersCategCod
    Else
        lblCategoria = ""
        lblCodCat = ""
    End If
    If gsOpeCod = CStr(gCGArendirViatSust2MN) Or gsOpeCod = CStr(gCGArendirViatSust2ME) Then
        devolverARendirViaticosParaSustentar (user.PersCod)
    ElseIf gsOpeCod = CStr(gCGArendirCtatSust2MN) Or gsOpeCod = CStr(gCGArendirCtatSust2ME) Then
        devolverARendirCuentasParaSustentar (user.PersCod)
    End If
    devolverDatosColaborador = True
Else
    Exit Function
End If

End Function

Private Sub devolverARendirViaticosParaSustentar(ByVal psPersCod As String)
Dim oNArendir As NARendir
Set oNArendir = New NARendir
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset

Set rsSustentar = oNArendir.obtenerARendirViaticosParaSustentar(IIf(Mid(gsOpeCod, 3, 1) = "1", gMonedaNacional, gMonedaExtranjera), psPersCod)
Call LimpiaFlex(FEAtenciones)

If Not rsSustentar.BOF And Not rsSustentar.EOF Then
    FEAtenciones.lbEditarFlex = True
    Do While Not rsSustentar.EOF
        FEAtenciones.AdicionaFila
        FEAtenciones.TextMatrix(FEAtenciones.row, 1) = rsSustentar!cDocAbrev
        FEAtenciones.TextMatrix(FEAtenciones.row, 2) = rsSustentar!cDocNro
        FEAtenciones.TextMatrix(FEAtenciones.row, 3) = rsSustentar!cDocFecha
        FEAtenciones.TextMatrix(FEAtenciones.row, 4) = Format(rsSustentar!nMontoAtendido, gsFormatoNumeroView)
        FEAtenciones.TextMatrix(FEAtenciones.row, 5) = Format(rsSustentar!nMovSaldo, gsFormatoNumeroView)
        FEAtenciones.TextMatrix(FEAtenciones.row, 6) = rsSustentar!cMovDesc
        FEAtenciones.TextMatrix(FEAtenciones.row, 7) = rsSustentar!nMovNro
        FEAtenciones.TextMatrix(FEAtenciones.row, 8) = rsSustentar!cDocDesc
        rsSustentar.MoveNext
    Loop
    FEAtenciones.lbEditarFlex = False
Else
    MsgBox "No tiene Viaticos para sustentar", vbInformation, "Aviso"
    limpiarCampos
End If

Set oNArendir = Nothing
Set rsSustentar = Nothing
End Sub

Private Sub devolverARendirCuentasParaSustentar(ByVal psPersCod As String)
Dim oNArendir As NARendir
Set oNArendir = New NARendir
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset

Set rsSustentar = oNArendir.obtenerARendirCuentasParaSustentar(IIf(Mid(gsOpeCod, 3, 1) = "1", gMonedaNacional, gMonedaExtranjera), psPersCod)
Call LimpiaFlex(FEAtenciones)

If Not rsSustentar.BOF And Not rsSustentar.EOF Then
    FEAtenciones.lbEditarFlex = True
    Do While Not rsSustentar.EOF
        FEAtenciones.AdicionaFila
        FEAtenciones.TextMatrix(FEAtenciones.row, 1) = rsSustentar!cDocAbrev
        FEAtenciones.TextMatrix(FEAtenciones.row, 2) = rsSustentar!cDocNro
        FEAtenciones.TextMatrix(FEAtenciones.row, 3) = rsSustentar!cDocFecha
        FEAtenciones.TextMatrix(FEAtenciones.row, 4) = Format(rsSustentar!nMontoAtendido, gsFormatoNumeroView)
        FEAtenciones.TextMatrix(FEAtenciones.row, 5) = Format(rsSustentar!nMovSaldo, gsFormatoNumeroView)
        FEAtenciones.TextMatrix(FEAtenciones.row, 6) = rsSustentar!cMovDesc
        FEAtenciones.TextMatrix(FEAtenciones.row, 7) = rsSustentar!nMovNro
        FEAtenciones.TextMatrix(FEAtenciones.row, 8) = rsSustentar!cDocDesc
        FEAtenciones.TextMatrix(FEAtenciones.row, 9) = rsSustentar!nMovNroAte
        rsSustentar.MoveNext
    Loop
    FEAtenciones.lbEditarFlex = False
Else
    MsgBox "No tiene Viaticos para sustentar", vbInformation, "Aviso"
    limpiarCampos
End If

Set oNArendir = Nothing
Set rsSustentar = Nothing
End Sub

Private Function verificarLimiteSustentar(ByVal psRHCargoCategoria As String, ByVal pnViaticoMovNro As Long, ByVal pnTipoArendir As Integer) As Boolean
Dim oNArendir As NARendir
Set oNArendir = New NARendir
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
Dim rsFecha As ADODB.Recordset
Set rsFecha = New ADODB.Recordset
Dim lnDiasLimite, i, j As Integer
Dim X, Y As Integer '************ Agregado por PASI20131119 segun TI-ERS107-2013
Dim ldFechaLlegada, ldFechaEnInstitucion, ldFechaLimite As Date


verificarLimiteSustentar = False

ldFechaLlegada = oNArendir.obtenerFechaLlegadaColaborador(pnViaticoMovNro, pnTipoArendir)
lnDiasLimite = oNArendir.obtenerLimiteSustentarCategoria(psRHCargoCategoria, pnTipoArendir)

If ldFechaLlegada <> "01/01/1900" Then
    'Modificado PASIERS0172015
'    I = 1
'    Do While oNContFunciones.obtenerSiFechaEsLaborable(DateAdd("D", I, ldFechaLlegada), gsCodAge) = False
'        I = I + 1
'    Loop
'    If pnTipoArendir = gArendirTipoViaticos Then
'        ldFechaEnInstitucion = DateAdd("D", I, ldFechaLlegada)
'    Else
'        ldFechaEnInstitucion = ldFechaLlegada
'    End If
    ldFechaEnInstitucion = ldFechaLlegada
    'end PASI
Else
  Exit Function
End If

'**************************************************************************************
'Modificado por PASI20131119 segun TI-ERS107-2013
'If Trim(psRHCargoCategoria) = "" Then
'    j = 3
'Else
'    j = lnDiasLimite
'End If
'Do While oNContFunciones.obtenerSiFechaEsLaborable(DateAdd("D", j, ldFechaEnInstitucion), gsCodAge) = False
'   j = j + 1
'Loop
'ldFechaLimite = DateAdd("D", j, ldFechaEnInstitucion)

If Trim(psRHCargoCategoria) = "" Then
    If pnTipoArendir = gArendirTipoViaticos Then
        j = 5
    Else
        j = 3
    End If
Else
    j = lnDiasLimite
End If

X = 0
Y = 0

    Do While X < j
        If oNContFunciones.obtenerSiFechaEsLaborable(DateAdd("D", Y + 1, ldFechaEnInstitucion), gsCodAge) = True Then
            X = X + 1
        End If
        Y = Y + 1
    Loop
ldFechaLimite = DateAdd("D", Y, ldFechaEnInstitucion)
'END PASI****************************************************************************

If ldFechaEnInstitucion >= gdFecSis Or ldFechaLimite >= gdFecSis Then
    verificarLimiteSustentar = True
End If



End Function

Private Function verificarProrrogaSustentar() As Integer
Dim oNArendir As NARendir
Set oNArendir = New NARendir
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
Dim ldFechaProrroga As Date
Dim ldFechaLimite As Date 'PASI20140507 TI-ERS060-2014
Dim i As Integer
Dim X, Y As Integer '***Agregado por PASI20140102 TI-ERS107-2013

Set rsSustentar = oNArendir.devolverProrrogaSustentar(CLng(FEAtenciones.TextMatrix(FEAtenciones.row, 7)))

If Not rsSustentar.BOF And Not rsSustentar.EOF Then
    
    'Comentado PASI20140507 TI-ERS060-2014
    'I = rsSustentar!ndias
    i = rsSustentar!nNroProrroga * rsSustentar!nDias 'Modificado PASI20140507 TI-ERS060-2014
    'end PASI
    
    '***Modificado por PASI20140102 TI-ERS107-2013
    'Do While oNContFunciones.obtenerSiFechaEsLaborable(DateAdd("D", i, CDate(rsSustentar!cFecha)), gsCodArea) = False
    '    i = i + 1
    'Loop
    'ldFechaProrroga = DateAdd("D", i, CDate(rsSustentar!cFecha))
    
    'Modificado PASI20140507 TI-ERS060-2014
'    Do While X < I
'        If oNContFunciones.obtenerSiFechaEsLaborable(DateAdd("D", Y + 1, CDate(rsSustentar!cFecha)), gsCodArea) = True Then
'            X = X + 1
'        End If
'        Y = Y + 1
'    Loop
'    ldFechaProrroga = DateAdd("D", Y, CDate(rsSustentar!cFecha))
     '***Fin PASI
     
    ldFechaLimite = FechaLimiteSustentar(lblCodCat, lnMovNro, lnTipoArendir)
    Do While X < i
        'If oNContFunciones.obtenerSiFechaEsLaborable(DateAdd("D", Y + 1, ldFechaLimite), gsCodArea) = True Then
        If oNContFunciones.obtenerSiFechaEsLaborable(DateAdd("D", Y + 1, ldFechaLimite), gsCodAge) = True Then 'PASI20150730
            X = X + 1
        End If
        Y = Y + 1
    Loop
    ldFechaProrroga = DateAdd("D", Y, CDate(ldFechaLimite))
'End PASI
    
    If ldFechaProrroga >= gdFecSis Then
        verificarProrrogaSustentar = 1
    ElseIf ldFechaProrroga < gdFecSis Then
        verificarProrrogaSustentar = 2
    End If
Else
    verificarProrrogaSustentar = 3
End If

ldFechaProrroga = "01/01/1900"
i = 0
Set oNArendir = Nothing
Set rsSustentar = Nothing
Set oNContFunciones = Nothing

End Function

Function validarDatosColaborador() As Boolean
validarDatosColaborador = False
    If FEAtenciones.TextMatrix(1, 0) = "" Then
        MsgBox "No hay Viáticos para sustentar", vbInformation, "Aviso"
        Exit Function
    End If

    If txtBuscaPers = "" Then
        MsgBox "Falta ingresar el código de un Colaborador", vbInformation, "Aviso"
        Exit Function
    End If

    If txtBuscaPers <> txtBuscaPers.psCodigoPersona Then
        MsgBox "El código ingresado no coindide con el del Colaborador visualizado", vbInformation, "Aviso"
        Exit Function
    End If
validarDatosColaborador = True
End Function
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSustentar_Click()
Dim oFrmOpeRegDocs As frmOpeRegDocs
Set oFrmOpeRegDocs = New frmOpeRegDocs
Dim lsNroArendir, lsNroDoc, lsFechaDoc, lsPersCod As String
Dim lsPersNomb, lsAreaCod, lsAreaDesc, lsAgeCod As String
Dim lsAgeDesc, lsDescDoc, lsMovNroAtenc  As String
Dim lsCtaArendir, lsCtaPendiente, lsMovNroSolicitud As String
Dim lnImporte, lnSaldo As Currency
Dim nProrroga As Integer


If validarDatosColaborador = False Then
    Exit Sub
End If

lnMovNro = FEAtenciones.TextMatrix(FEAtenciones.row, 7)  'Agregado PASI20140507 TI-ERS060-2014
nProrroga = verificarProrrogaSustentar

If nProrroga = 2 Then
    '*** Modificado por PASI20140102 TI-ERS107-2013
    'MsgBox "La fecha de prorroga para sustentar ya fue superano. Consultar a Área de Contabilidad", vbInformation, "Aviso"
    MsgBox "La fecha de prorroga para sustentar ya fue superado. Consultar a Área de Contabilidad", vbInformation, "Aviso"
    Exit Sub
    '***FIN PASI
ElseIf nProrroga = 3 Then
    
    If verificarLimiteSustentar(lblCodCat, lnMovNro, lnTipoArendir) = False Then 'Modificado PASI20140507 TI-ERS060-2014 ; cambiado FEAtenciones.TextMatrix(FEAtenciones.row, 7) por lnMovNro
        MsgBox "Plazo de sustentación vencido." & Chr(13) & "Consultar Reglamento de Entregas a Rendir en Intranet", vbInformation, "Aviso"
        Exit Sub
    End If
End If



lsNroArendir = FEAtenciones.TextMatrix(FEAtenciones.row, 2)
lsNroDoc = FEAtenciones.TextMatrix(FEAtenciones.row, 2)
lsFechaDoc = FEAtenciones.TextMatrix(FEAtenciones.row, 3)
lsPersCod = txtBuscaPers.psCodigoPersona
lsPersNomb = lblPersNombre
lsAreaCod = lblAreaCod
lsAreaDesc = lblAreaDesc
lsAgeCod = lblAgeCod
lsAgeDesc = lblAgeDesc
lsDescDoc = FEAtenciones.TextMatrix(FEAtenciones.row, 8)
If gsOpeCod = CStr(gCGArendirViatSust2MN) Or gsOpeCod = CStr(gCGArendirViatSust2ME) Then
    lsMovNroAtenc = FEAtenciones.TextMatrix(FEAtenciones.row, 7)
ElseIf gsOpeCod = CStr(gCGArendirCtatSust2MN) Or gsOpeCod = CStr(gCGArendirCtatSust2ME) Then
    lsMovNroAtenc = FEAtenciones.TextMatrix(FEAtenciones.row, 9)
End If
lnImporte = FEAtenciones.TextMatrix(FEAtenciones.row, 4)
lnSaldo = FEAtenciones.TextMatrix(FEAtenciones.row, 5)
lsMovNroSolicitud = FEAtenciones.TextMatrix(FEAtenciones.row, 7)

oFrmOpeRegDocs.Inicio lnArendirFase, lnTipoArendir, False, lsNroArendir, _
                      lsNroDoc, lsFechaDoc, lsPersCod, lsPersNomb, _
                      lsAreaCod, lsAreaDesc, lsAgeCod, lsAgeDesc, _
                      lsDescDoc, lsMovNroAtenc, lnImporte, fsCtaArendir, _
                      fsCtaPendiente, lnSaldo, lsMovNroSolicitud
                      
FEAtenciones.TextMatrix(FEAtenciones.row, 5) = Format(oFrmOpeRegDocs.lnSaldo, gsFormatoNumeroView)
FEAtenciones.SetFocus
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Sustento :" & lblPersNombre
            Set objPista = Nothing
            '*******
End Sub

Private Sub FEAtenciones_Click()
    Call actualizarAtencion
End Sub

Private Sub FEAtenciones_GotFocus()
    Call actualizarAtencion
End Sub

Private Sub FEAtenciones_OnRowChange(pnRow As Long, pnCol As Long)
    Call actualizarAtencion
End Sub
Private Sub FEAtenciones_RowColChange()
    Call actualizarAtencion
End Sub

Private Sub Form_Activate()
If fnSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
    Dim oDOperacion As DOperacion
    Set oDOperacion = New DOperacion
    
    CentraForm Me
    
    devolverDatosColaborador "", gsCodPersUser
    
    fsCtaArendir = oDOperacion.EmiteOpeCta(gsOpeCod, "H", "0")
        If fsCtaArendir = "" Then
        MsgBox "Falta asignar Cuenta a Rendir" & oImpresora.gPrnSaltoLinea & "Por favor consultar con Contabilidad", vbInformation, "Aviso"
        fnSalir = True
        Exit Sub
    End If
    
    fsCtaPendiente = oDOperacion.EmiteOpeCta(gsOpeCod, "H", "1")
    If fsCtaPendiente = "" Then
        MsgBox "Falta asignar Cuenta de Pendiente a Operación" & oImpresora.gPrnSaltoLinea & "Por favor consultar con Contabilidad", vbInformation, "Aviso"
        fnSalir = True
        Exit Sub
    End If
    Set oDOperacion = Nothing
End Sub

Private Sub txtBuscaPers_EmiteDatos()
    If devolverDatosColaborador("", txtBuscaPers) = False Then
        If txtBuscaPers <> "" Then
            MsgBox "Persona ingresada no se encuentra registrada como empleado de la Institucion", vbInformation, "Aviso"
            limpiarCampos
            Exit Sub
        End If
    End If
    FEAtenciones.SetFocus
End Sub

'Agregado PASI20140507 TI-ERS060 2014
Private Function FechaLimiteSustentar(ByVal psRHCargoCategoria As String, ByVal pnViaticoMovNro As Long, ByVal pnTipoArendir As Integer) As Date
Dim oNArendir As NARendir
Set oNArendir = New NARendir
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
Dim rsFecha As ADODB.Recordset
Set rsFecha = New ADODB.Recordset
Dim lnDiasLimite, i, j As Integer
Dim X, Y As Integer
Dim ldFechaLlegada, ldFechaEnInstitucion, ldFechaLimite As Date


'verificarLimiteSustentar = False

ldFechaLlegada = oNArendir.obtenerFechaLlegadaColaborador(pnViaticoMovNro, pnTipoArendir)
lnDiasLimite = oNArendir.obtenerLimiteSustentarCategoria(psRHCargoCategoria, pnTipoArendir)

If ldFechaLlegada <> "01/01/1900" Then
    'Modificado PASIERS0172015
'    I = 1
'    Do While oNContFunciones.obtenerSiFechaEsLaborable(DateAdd("D", I, ldFechaLlegada), gsCodAge) = False
'        I = I + 1
'    Loop
'    If pnTipoArendir = gArendirTipoViaticos Then
'        ldFechaEnInstitucion = DateAdd("D", I, ldFechaLlegada)
'    Else
'        ldFechaEnInstitucion = ldFechaLlegada
'    End If
    ldFechaEnInstitucion = ldFechaLlegada
Else
  Exit Function
End If

If Trim(psRHCargoCategoria) = "" Then
    If pnTipoArendir = gArendirTipoViaticos Then
        j = 5
    Else
        j = 3
    End If
Else
    j = lnDiasLimite
End If

X = 0
Y = 0

    Do While X < j
        If oNContFunciones.obtenerSiFechaEsLaborable(DateAdd("D", Y + 1, ldFechaEnInstitucion), gsCodAge) = True Then
            X = X + 1
        End If
        Y = Y + 1
    Loop
    ldFechaLimite = DateAdd("D", Y, ldFechaEnInstitucion)
    FechaLimiteSustentar = ldFechaLimite
End Function
'end PASI
