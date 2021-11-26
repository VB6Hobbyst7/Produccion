VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHojaRutaAnalistaGeneraResultado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resultado de Hoja de Ruta"
   ClientHeight    =   6330
   ClientLeft      =   6975
   ClientTop       =   5025
   ClientWidth     =   16005
   ControlBox      =   0   'False
   Icon            =   "frmHojaRutaAnalistaGeneraResultado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   16005
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   12000
      TabIndex        =   7
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmbCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   14040
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmbVisitasNoPlaneadas 
      Caption         =   "Visitas No planeadas"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmbRegistrarResultado 
      Caption         =   "Registrar Resultado"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   1815
   End
   Begin TabDlg.SSTab ssTabMain 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Clientes Promocionales"
      TabPicture(0)   =   "frmHojaRutaAnalistaGeneraResultado.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxClientesPromocionales"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Clientes en Mora"
      TabPicture(1)   =   "frmHojaRutaAnalistaGeneraResultado.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxClientesMora"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Web, Kiosko y Facebook"
      TabPicture(2)   =   "frmHojaRutaAnalistaGeneraResultado.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flxClientesWeb"
      Tab(2).ControlCount=   1
      Begin SICMACT.FlexEdit flxClientesMora 
         Height          =   4695
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   8281
         Cols0           =   14
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmHojaRutaAnalistaGeneraResultado.frx":035E
         EncabezadosAnchos=   "600-1500-2500-2000-2000-1200-1200-1000-1200-1200-1200-1200-1200-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-L-L-R-C-R-C-R-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-3-5-2-2-3-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit flxClientesPromocionales 
         Height          =   4695
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   8281
         Cols0           =   14
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmHojaRutaAnalistaGeneraResultado.frx":03FB
         EncabezadosAnchos=   "600-1500-2500-1200-2500-2000-2000-1200-1200-1200-1200-1000-1200-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-L-L-L-C-C-R-C-R-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-2-2-3-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit flxClientesWeb 
         Height          =   4695
         Left            =   -74760
         TabIndex        =   6
         Top             =   600
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   8281
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cliente-DOI-Dirección-Profesión-Teléfono-Móvil-Monto Solicitado-N°Cuotas-Observaciones-nLineaRutaId"
         EncabezadosAnchos=   "600-2500-1200-2500-2000-1200-1200-2000-1000-1200-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-C-C-R-R-L-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-2-3-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmHojaRutaAnalistaGeneraResultado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dHojaRuta As New DCOMhojaRuta
Dim rsClientesPromocionales As ADODB.Recordset
Dim rsClientesMora As ADODB.Recordset
Dim nMora As Integer
Dim nPromo As Integer
Dim bNoCerrarAlsalir As Boolean

Public Sub inicio()
    Dim cPeriodo As String: cPeriodo = Format(gdFecSis, "YYYYMM")
    Dim oDhoja As DCOMhojaRuta
    Set oDhoja = dHojaRuta
    bNoCerrarAlsalir = True
    If oDhoja.haConfiguradoAgencia(cPeriodo, gsCodAge) Then
        Dim nPendiestesAtrasados As Integer: nPendiestesAtrasados = oDhoja.obtenerNumeroVisitasPendientes(gsCodUser, 0)
        If nPendiestesAtrasados = 0 Then
            Me.Show 1
        Else
            'ACA DEBE SOLICITAR EL VISTO DEL JEFE DE AGENCIA
            If Not oDhoja.tieneVistoPendiente(gsCodUser) Then
                Dim cMovNro As String: cMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                oDhoja.solicitarVisto gsCodUser, cMovNro
            End If
            If frmHojaRutaAnalistaVistoInc.inicio Then
                Me.Show 1
            Else
                Unload Me
            End If
        End If
    Else
        MsgBox "Aún no existe Configuración de Hoja de Ruta para la Agencia, comuníquelo al Jefe de Agencia."
        Unload Me
    End If
End Sub

Private Sub cmbCerrar_Click()
    If nMora + nPromo > 0 Then
        If bNoCerrarAlsalir Then
            Unload Me
            Exit Sub
        End If
        Dim resp As String: resp = MsgBox("Aun no ha completado sus resultados, ¿desea salir del sistema?", vbYesNo, "Confirmar")
        If resp = vbYes Then
            End
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmbVisitasNoPlaneadas_Click()
    If ssTabMain.Tab = 0 Or ssTabMain.Tab = 2 Then  'visitas promocion
        Call frmHojaRutaAnalistaResPromo.inicio("", 0, False)
    Else
        Call frmHojaRutaAnalistaResMora.inicio("", 0, False)
    End If
End Sub

'WIOR 20151125 ***
Private Sub cmdActualizar_Click()
    Call Form_Load
End Sub
'WIOR FIN ********

Private Sub Form_Load()
    'cargar la lista de clientes
    llenarClientesPromocionales
    llenarClientesMora
    ssTabMain.Tab = 0
End Sub
Private Function llenarClientesPromocionales()
    nPromo = 0
    Set rsClientesPromocionales = dHojaRuta.ObtenerHojaRutaDia(gsCodUser, 1)
    
    Dim nRow As Integer
    Dim nRow2 As Integer
    
    LimpiaFlex flxClientesPromocionales
    flxClientesPromocionales.Rows = 2
    
    LimpiaFlex flxClientesWeb
    flxClientesWeb.Rows = 2
    
    Do While Not rsClientesPromocionales.EOF
    
        If rsClientesPromocionales!bWeb = 0 Then
            flxClientesPromocionales.AdicionaFila
            nRow = flxClientesPromocionales.Rows - 1
            flxClientesPromocionales.TextMatrix(nRow, 1) = rsClientesPromocionales!cPersCodCliente
            flxClientesPromocionales.TextMatrix(nRow, 2) = rsClientesPromocionales!cPersNombre
            flxClientesPromocionales.TextMatrix(nRow, 3) = rsClientesPromocionales!cDOI
            flxClientesPromocionales.TextMatrix(nRow, 4) = rsClientesPromocionales!cDireccion
            flxClientesPromocionales.TextMatrix(nRow, 5) = rsClientesPromocionales!cZona
            flxClientesPromocionales.TextMatrix(nRow, 6) = rsClientesPromocionales!cActividad
            flxClientesPromocionales.TextMatrix(nRow, 7) = rsClientesPromocionales!cTelefono
            flxClientesPromocionales.TextMatrix(nRow, 8) = rsClientesPromocionales!cMovil
            flxClientesPromocionales.TextMatrix(nRow, 9) = rsClientesPromocionales!nMontoUltCred
            flxClientesPromocionales.TextMatrix(nRow, 10) = rsClientesPromocionales!nEndeudamiento
            flxClientesPromocionales.TextMatrix(nRow, 11) = rsClientesPromocionales!nNIFIS
            flxClientesPromocionales.TextMatrix(nRow, 12) = rsClientesPromocionales!cObservaciones
            flxClientesPromocionales.TextMatrix(nRow, 13) = rsClientesPromocionales!nLineaRutaId
        Else
            flxClientesWeb.AdicionaFila
            nRow2 = flxClientesWeb.Rows - 1
            flxClientesWeb.TextMatrix(nRow2, 1) = rsClientesPromocionales!cPersNombre
            flxClientesWeb.TextMatrix(nRow2, 2) = rsClientesPromocionales!cDOI
            flxClientesWeb.TextMatrix(nRow2, 3) = rsClientesPromocionales!cDireccion
            flxClientesWeb.TextMatrix(nRow2, 4) = rsClientesPromocionales!cActividad
            flxClientesWeb.TextMatrix(nRow2, 5) = rsClientesPromocionales!cTelefono
            flxClientesWeb.TextMatrix(nRow2, 6) = rsClientesPromocionales!cMovil
            flxClientesWeb.TextMatrix(nRow2, 7) = rsClientesPromocionales!nMontoUltCred
            flxClientesWeb.TextMatrix(nRow2, 8) = rsClientesPromocionales!nNIFIS
            flxClientesWeb.TextMatrix(nRow2, 9) = rsClientesPromocionales!cObservaciones
            flxClientesWeb.TextMatrix(nRow2, 10) = rsClientesPromocionales!nLineaRutaId
            
        End If
        
        rsClientesPromocionales.MoveNext
        nPromo = nPromo + 1
    Loop
End Function
Private Function llenarClientesMora()
    nMora = 0
    Set rsClientesMora = dHojaRuta.ObtenerHojaRutaDia(gsCodUser, 0)
    Dim nRow As Integer
    LimpiaFlex flxClientesMora
    flxClientesMora.Rows = 2
    Do While Not rsClientesMora.EOF
        flxClientesMora.AdicionaFila
        nRow = flxClientesMora.Rows - 1
        flxClientesMora.TextMatrix(nRow, 1) = rsClientesMora!cPersCodCliente
        flxClientesMora.TextMatrix(nRow, 2) = rsClientesMora!cPersNombre
        flxClientesMora.TextMatrix(nRow, 3) = rsClientesMora!cDireccion
        flxClientesMora.TextMatrix(nRow, 4) = rsClientesMora!cDireccionNegocio
        flxClientesMora.TextMatrix(nRow, 5) = rsClientesMora!cTelefono
        flxClientesMora.TextMatrix(nRow, 6) = rsClientesMora!cMovil
        flxClientesMora.TextMatrix(nRow, 7) = rsClientesMora!nCuota
        flxClientesMora.TextMatrix(nRow, 8) = rsClientesMora!dFechaVenc
        flxClientesMora.TextMatrix(nRow, 9) = rsClientesMora!nMontoCuota
        flxClientesMora.TextMatrix(nRow, 10) = rsClientesMora!nMora
        flxClientesMora.TextMatrix(nRow, 11) = rsClientesMora!nDiasAtraso
        flxClientesMora.TextMatrix(nRow, 12) = rsClientesMora!cObservaciones
        flxClientesMora.TextMatrix(nRow, 13) = rsClientesMora!nLineaRutaId
        rsClientesMora.MoveNext
        nMora = nMora + 1
    Loop
End Function

Private Sub cmbRegistrarResultado_Click()
    Dim cPersCod As String
    Dim cPersNom As String
    Dim nLineaRutaId As Integer
    If ssTabMain.Tab = 0 Or ssTabMain.Tab = 2 Then  'visitas promocion
    
        If ssTabMain.Tab = 0 Then
            cPersCod = flxClientesPromocionales.TextMatrix(flxClientesPromocionales.row, 1)
            cPersNom = flxClientesPromocionales.TextMatrix(flxClientesPromocionales.row, 2)
        
            If flxClientesPromocionales.TextMatrix(flxClientesPromocionales.row, 13) = "" Then
                MsgBox "Debe elegir un cliente"
                Exit Sub
            End If
            nLineaRutaId = flxClientesPromocionales.TextMatrix(flxClientesPromocionales.row, 13)
            
            If frmHojaRutaAnalistaResPromo.inicio(cPersCod, nLineaRutaId, True, cPersNom) Then
                flxClientesPromocionales.EliminaFila flxClientesPromocionales.row, True
                nPromo = nPromo - 1
            End If
        Else
            If flxClientesWeb.TextMatrix(flxClientesWeb.row, 10) = "" Then
                MsgBox "Debe elegir un cliente"
                Exit Sub
            End If
            cPersCod = ""
            cPersNom = flxClientesWeb.TextMatrix(flxClientesWeb.row, 1)
            nLineaRutaId = flxClientesWeb.TextMatrix(flxClientesWeb.row, 10)
            
            If frmHojaRutaAnalistaResPromo.inicio(cPersCod, nLineaRutaId, True, cPersNom) Then
                flxClientesWeb.EliminaFila flxClientesWeb.row, True
                nPromo = nPromo - 1
            End If
        End If
    Else 'visitas mora
        
        If flxClientesMora.TextMatrix(flxClientesMora.row, 13) = "" Then
                MsgBox "Debe elegir un cliente"
                Exit Sub
        End If
        
        cPersCod = flxClientesMora.TextMatrix(flxClientesMora.row, 1)
        nLineaRutaId = flxClientesMora.TextMatrix(flxClientesMora.row, 13)
        cPersNom = flxClientesMora.TextMatrix(flxClientesMora.row, 2)
        
        If frmHojaRutaAnalistaResMora.inicio(cPersCod, nLineaRutaId, True, cPersNom) Then
            flxClientesMora.EliminaFila flxClientesMora.row, True
            nMora = nMora - 1
        End If
        
    End If
    comprobarTermino
    'llenarClientesMora
    'llenarClientesPromocionales
End Sub

Private Sub comprobarTermino()
    If nMora + nPromo <= 0 Then
        Dim resp As String: resp = MsgBox("Ud ha terminado de registrar sus visitas, desea cerrar este formulario", vbYesNo, "Confirmar")
        If resp = vbYes Then Unload Me
    End If
End Sub

